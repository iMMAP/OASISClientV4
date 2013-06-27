VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddIncident 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incident"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8610
   Icon            =   "frmAddIncident.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elAddPt 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8610
      _cx             =   15187
      _cy             =   9340
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
      Align           =   5
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
      _GridInfo       =   $"frmAddIncident.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic elButtons 
         Height          =   510
         Left            =   30
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4755
         Width           =   8550
         _cx             =   15081
         _cy             =   900
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
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3645
            TabIndex        =   17
            Top             =   0
            Width           =   1140
         End
         Begin VB.CommandButton cmdSubmit 
            Caption         =   "Submit"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2475
            TabIndex        =   16
            Top             =   0
            Width           =   1140
         End
         Begin VB.CommandButton cmdForward 
            Caption         =   "Next >>"
            Height          =   375
            Left            =   1290
            TabIndex        =   15
            Top             =   0
            Width           =   1140
         End
         Begin VB.CommandButton cmdBack 
            Caption         =   "<< Back"
            Height          =   375
            Left            =   135
            TabIndex        =   14
            Top             =   0
            Width           =   1140
         End
      End
      Begin C1SizerLibCtl.C1Tab c1TabAddPt 
         Height          =   4695
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8550
         _cx             =   15081
         _cy             =   8281
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
         Caption         =   "Start|General|Location|Time|Type|Attachments|Summary"
         Align           =   0
         CurrTab         =   2
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
         Flags(5)        =   2
         Begin C1SizerLibCtl.C1Elastic elSummary 
            Height          =   4365
            Left            =   10065
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   315
            Width           =   8520
            _cx             =   15028
            _cy             =   7699
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
            Begin C1SizerLibCtl.C1Elastic elMainSummary 
               Height          =   4365
               Left            =   0
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   0
               Width           =   8520
               _cx             =   15028
               _cy             =   7699
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
               BackColor       =   16777215
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   ""
               Align           =   5
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
               FrameWidth      =   0
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   ""
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.CommandButton cmdAddAttachments 
                  Caption         =   "Attachments"
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   163
                  Top             =   3960
                  Width           =   1095
               End
               Begin VB.CheckBox chkCreateEmail 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Attach To Email"
                  Height          =   255
                  Left            =   180
                  TabIndex        =   155
                  Top             =   4350
                  Visible         =   0   'False
                  Width           =   1695
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Report Formats:"
                  Height          =   765
                  Left            =   90
                  TabIndex        =   152
                  Top             =   3540
                  Width           =   2655
                  Begin VB.CheckBox chkOASISReports 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Export to OASIS Reports"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   154
                     Top             =   180
                     Value           =   1  'Checked
                     Width           =   2265
                  End
                  Begin VB.CheckBox chkHTMLReport 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Export to HTML Document"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   153
                     Top             =   450
                     Value           =   1  'Checked
                     Width           =   2295
                  End
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   0
                  Left            =   3465
                  TabIndex        =   109
                  Top             =   45
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   1
                  Left            =   3465
                  TabIndex        =   108
                  Top             =   360
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   2
                  Left            =   3465
                  TabIndex        =   107
                  Top             =   675
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   3
                  Left            =   3465
                  TabIndex        =   106
                  Top             =   990
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   4
                  Left            =   3465
                  TabIndex        =   105
                  Top             =   1305
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   5
                  Left            =   3465
                  TabIndex        =   104
                  Top             =   1620
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   6
                  Left            =   3465
                  TabIndex        =   103
                  Top             =   1935
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   7
                  Left            =   3465
                  TabIndex        =   102
                  Top             =   2250
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   8
                  Left            =   3465
                  TabIndex        =   101
                  Top             =   2565
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   9
                  Left            =   3465
                  TabIndex        =   100
                  Top             =   2880
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   10
                  Left            =   3465
                  TabIndex        =   99
                  Top             =   3195
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   11
                  Left            =   7245
                  TabIndex        =   98
                  Top             =   45
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   12
                  Left            =   7245
                  TabIndex        =   97
                  Top             =   360
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   13
                  Left            =   7245
                  TabIndex        =   96
                  Top             =   675
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   14
                  Left            =   7245
                  TabIndex        =   95
                  Top             =   990
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   15
                  Left            =   7245
                  TabIndex        =   94
                  Top             =   1305
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   16
                  Left            =   7245
                  TabIndex        =   93
                  Top             =   1620
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   17
                  Left            =   7245
                  TabIndex        =   92
                  Top             =   1935
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   18
                  Left            =   7245
                  TabIndex        =   91
                  Top             =   2250
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   19
                  Left            =   7245
                  TabIndex        =   90
                  Top             =   2565
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   20
                  Left            =   7245
                  TabIndex        =   89
                  Top             =   2880
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   21
                  Left            =   7245
                  TabIndex        =   88
                  Top             =   3195
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   22
                  Left            =   7245
                  TabIndex        =   87
                  Top             =   3825
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   23
                  Left            =   7245
                  TabIndex        =   86
                  Top             =   4140
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   24
                  Left            =   7245
                  TabIndex        =   85
                  Top             =   4455
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   25
                  Left            =   7245
                  TabIndex        =   84
                  Top             =   4770
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   26
                  Left            =   7245
                  TabIndex        =   83
                  Top             =   5085
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   27
                  Left            =   7245
                  TabIndex        =   82
                  Top             =   5400
                  Visible         =   0   'False
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   28
                  Left            =   7245
                  TabIndex        =   81
                  Top             =   5715
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   29
                  Left            =   7245
                  TabIndex        =   80
                  Top             =   6030
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   30
                  Left            =   7245
                  TabIndex        =   79
                  Top             =   6345
                  Width           =   330
               End
               Begin VB.CommandButton cmdWizValUpdate 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   31
                  Left            =   7245
                  TabIndex        =   78
                  Top             =   6660
                  Width           =   330
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   0
                  Left            =   45
                  TabIndex        =   141
                  Top             =   90
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   1
                  Left            =   45
                  TabIndex        =   140
                  Top             =   405
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   2
                  Left            =   45
                  TabIndex        =   139
                  Top             =   720
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   3
                  Left            =   45
                  TabIndex        =   138
                  Top             =   1035
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   4
                  Left            =   45
                  TabIndex        =   137
                  Top             =   1350
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   5
                  Left            =   45
                  TabIndex        =   136
                  Top             =   1665
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   6
                  Left            =   45
                  TabIndex        =   135
                  Top             =   1980
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   7
                  Left            =   45
                  TabIndex        =   134
                  Top             =   2295
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   8
                  Left            =   45
                  TabIndex        =   133
                  Top             =   2610
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   9
                  Left            =   45
                  TabIndex        =   132
                  Top             =   2925
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   10
                  Left            =   45
                  TabIndex        =   131
                  Top             =   3240
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   11
                  Left            =   3825
                  TabIndex        =   130
                  Top             =   90
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   12
                  Left            =   3825
                  TabIndex        =   129
                  Top             =   405
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   13
                  Left            =   3825
                  TabIndex        =   128
                  Top             =   720
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   14
                  Left            =   3825
                  TabIndex        =   127
                  Top             =   1035
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   15
                  Left            =   3825
                  TabIndex        =   126
                  Top             =   1350
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   16
                  Left            =   3825
                  TabIndex        =   125
                  Top             =   1665
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   17
                  Left            =   3825
                  TabIndex        =   124
                  Top             =   1980
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   18
                  Left            =   3825
                  TabIndex        =   123
                  Top             =   2295
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   19
                  Left            =   3825
                  TabIndex        =   122
                  Top             =   2610
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   20
                  Left            =   3825
                  TabIndex        =   121
                  Top             =   2925
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   21
                  Left            =   3825
                  TabIndex        =   120
                  Top             =   3240
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   22
                  Left            =   3825
                  TabIndex        =   119
                  Top             =   3870
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   23
                  Left            =   3825
                  TabIndex        =   118
                  Top             =   4185
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   24
                  Left            =   3825
                  TabIndex        =   117
                  Top             =   4500
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   25
                  Left            =   3825
                  TabIndex        =   116
                  Top             =   4815
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   26
                  Left            =   3825
                  TabIndex        =   115
                  Top             =   5130
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   27
                  Left            =   3825
                  TabIndex        =   114
                  Top             =   5445
                  Visible         =   0   'False
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   28
                  Left            =   3825
                  TabIndex        =   113
                  Top             =   5760
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   29
                  Left            =   3825
                  TabIndex        =   112
                  Top             =   6075
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   30
                  Left            =   3825
                  TabIndex        =   111
                  Top             =   6390
                  Width           =   885
               End
               Begin VB.Label lblWizValLabel 
                  AutoSize        =   -1  'True
                  Caption         =   "WizValLabel"
                  Height          =   195
                  Index           =   31
                  Left            =   3825
                  TabIndex        =   110
                  Top             =   6705
                  Width           =   885
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elAttachments 
            Height          =   4365
            Left            =   9765
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   315
            Width           =   8520
            _cx             =   15028
            _cy             =   7699
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   4365
               Left            =   0
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   0
               Width           =   8520
               _cx             =   15028
               _cy             =   7699
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
               FrameWidth      =   0
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmAddIncident.frx":6895
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic C1EReportOptions 
                  Height          =   510
                  Left            =   0
                  TabIndex        =   66
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1200
                  _cx             =   2117
                  _cy             =   900
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
                  BackColor       =   16777215
                  ForeColor       =   -2147483630
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   " Report Options"
                  Align           =   0
                  AutoSizeChildren=   0
                  BorderWidth     =   0
                  ChildSpacing    =   0
                  Splitter        =   0   'False
                  FloodDirection  =   0
                  FloodPercent    =   0
                  CaptionPos      =   0
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
                  FrameWidth      =   0
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   ""
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin VB.CheckBox chkAllowOther 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Allow other users to contact me"
                     Height          =   330
                     Left            =   2745
                     TabIndex        =   67
                     Top             =   180
                     Width           =   3525
                  End
                  Begin VB.CheckBox chkSendMe 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Send verification when published"
                     Height          =   285
                     Left            =   90
                     TabIndex        =   68
                     Top             =   225
                     Width           =   3615
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1EAttachments 
                  Height          =   3885
                  Left            =   0
                  TabIndex        =   69
                  TabStop         =   0   'False
                  Top             =   510
                  Width           =   1200
                  _cx             =   2117
                  _cy             =   6853
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial Black"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   900
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
                  Picture         =   "frmAddIncident.frx":68D3
                  Caption         =   " Attachments"
                  Align           =   0
                  AutoSizeChildren=   8
                  BorderWidth     =   0
                  ChildSpacing    =   0
                  Splitter        =   0   'False
                  FloodDirection  =   0
                  FloodPercent    =   0
                  CaptionPos      =   3
                  WordWrap        =   -1  'True
                  MaxChildSize    =   0
                  MinChildSize    =   0
                  TagWidth        =   0
                  TagPosition     =   0
                  Style           =   0
                  TagSplit        =   2
                  PicturePos      =   3
                  CaptionStyle    =   0
                  ResizeFonts     =   0   'False
                  GridRows        =   3
                  GridCols        =   1
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   0
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"frmAddIncident.frx":68EF
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin C1SizerLibCtl.C1Elastic C1ListPanel 
                     Height          =   3480
                     Left            =   0
                     TabIndex        =   70
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   1200
                     _cx             =   2117
                     _cy             =   6138
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
                     GridRows        =   2
                     GridCols        =   1
                     Frame           =   3
                     FrameStyle      =   0
                     FrameWidth      =   0
                     FrameColor      =   -2147483628
                     FrameShadow     =   -2147483632
                     FloodStyle      =   1
                     _GridInfo       =   $"frmAddIncident.frx":6932
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin VB.ListBox lstAttachments 
                        Height          =   3210
                        ItemData        =   "frmAddIncident.frx":6968
                        Left            =   0
                        List            =   "frmAddIncident.frx":696F
                        Style           =   1  'Checkbox
                        TabIndex        =   71
                        Top             =   255
                        Width           =   1200
                     End
                     Begin VB.Label lblAddedAttachments 
                        AutoSize        =   -1  'True
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "Added attachments"
                        Height          =   255
                        Left            =   0
                        TabIndex        =   72
                        Top             =   0
                        Width           =   1200
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                     Height          =   405
                     Left            =   0
                     TabIndex        =   73
                     TabStop         =   0   'False
                     Top             =   3480
                     Width           =   1200
                     _cx             =   2117
                     _cy             =   714
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
                     GridRows        =   1
                     GridCols        =   3
                     Frame           =   3
                     FrameStyle      =   0
                     FrameWidth      =   0
                     FrameColor      =   -2147483628
                     FrameShadow     =   -2147483632
                     FloodStyle      =   1
                     _GridInfo       =   $"frmAddIncident.frx":698D
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin VB.CommandButton cmdAdd 
                        Height          =   405
                        Left            =   840
                        Picture         =   "frmAddIncident.frx":69CB
                        Style           =   1  'Graphical
                        TabIndex        =   75
                        Top             =   0
                        Width           =   360
                     End
                     Begin VB.CommandButton cmdDelete 
                        Height          =   405
                        Left            =   975
                        Picture         =   "frmAddIncident.frx":D21D
                        Style           =   1  'Graphical
                        TabIndex        =   74
                        Top             =   0
                        Width           =   225
                     End
                  End
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elType 
            Height          =   4365
            Left            =   9465
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   315
            Width           =   8520
            _cx             =   15028
            _cy             =   7699
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
            Begin C1SizerLibCtl.C1Elastic C1Main 
               Height          =   4365
               Left            =   0
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   0
               Width           =   8520
               _cx             =   15028
               _cy             =   7699
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
               FrameWidth      =   0
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmAddIncident.frx":13A6F
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic C1RightPanel 
                  Height          =   4365
                  Left            =   4230
                  TabIndex        =   30
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   4290
                  _cx             =   7567
                  _cy             =   7699
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
                  FrameWidth      =   0
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"frmAddIncident.frx":13AAF
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin C1SizerLibCtl.C1Elastic C1EIncidentDescription 
                     Height          =   4365
                     Left            =   0
                     TabIndex        =   31
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   4290
                     _cx             =   7567
                     _cy             =   7699
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
                     BackColor       =   16777215
                     ForeColor       =   -2147483630
                     FloodColor      =   6553600
                     ForeColorDisabled=   -2147483631
                     Caption         =   "Description"
                     Align           =   0
                     AutoSizeChildren=   7
                     BorderWidth     =   6
                     ChildSpacing    =   0
                     Splitter        =   0   'False
                     FloodDirection  =   0
                     FloodPercent    =   0
                     CaptionPos      =   0
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
                     FrameWidth      =   0
                     FrameColor      =   -2147483628
                     FrameShadow     =   -2147483632
                     FloodStyle      =   1
                     _GridInfo       =   ""
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin VB.TextBox txtIncidentDescription 
                        Height          =   3900
                        Left            =   45
                        MultiLine       =   -1  'True
                        ScrollBars      =   2  'Vertical
                        TabIndex        =   32
                        Top             =   330
                        Width           =   4200
                     End
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   4365
                  Left            =   0
                  TabIndex        =   33
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   4230
                  _cx             =   7461
                  _cy             =   7699
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
                  GridRows        =   4
                  GridCols        =   1
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   0
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"frmAddIncident.frx":13AE1
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin VB.Frame frType 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Type of incident (choose the item that best apply)"
                     Height          =   675
                     Left            =   0
                     TabIndex        =   52
                     Top             =   675
                     Width           =   4230
                     Begin VB.ComboBox ComIncType 
                        Height          =   315
                        Left            =   90
                        TabIndex        =   53
                        Text            =   "Inc Type"
                        Top             =   270
                        Width           =   3975
                     End
                  End
                  Begin VB.Frame FraIncidentType 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "The incident you want to report was it?"
                     Height          =   675
                     Left            =   0
                     TabIndex        =   48
                     Top             =   0
                     Width           =   4230
                     Begin VB.OptionButton OptIncViolence 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "Unknown"
                        Height          =   285
                        Index           =   2
                        Left            =   2655
                        TabIndex        =   49
                        Top             =   270
                        Value           =   -1  'True
                        Width           =   1185
                     End
                     Begin VB.OptionButton OptIncViolence 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "Non violent"
                        Height          =   285
                        Index           =   1
                        Left            =   1305
                        TabIndex        =   50
                        Top             =   270
                        Width           =   1590
                     End
                     Begin VB.OptionButton OptIncViolence 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "Violent"
                        Height          =   285
                        Index           =   0
                        Left            =   135
                        TabIndex        =   51
                        Top             =   270
                        Width           =   1590
                     End
                  End
                  Begin VB.Frame FraSelectThe 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Select incident target (if not applicable choose N/A)"
                     Height          =   675
                     Left            =   0
                     TabIndex        =   46
                     Top             =   1350
                     Width           =   4230
                     Begin VB.ComboBox ComIncTarget 
                        Height          =   315
                        Left            =   90
                        TabIndex        =   47
                        Text            =   "Inc Target"
                        Top             =   225
                        Width           =   3930
                     End
                  End
                  Begin VB.Frame FrCas 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Were there any casualties?"
                     Height          =   2340
                     Left            =   0
                     TabIndex        =   34
                     Top             =   2025
                     Width           =   4230
                     Begin VB.Frame FraNumberOf 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "Number of casualties"
                        Enabled         =   0   'False
                        Height          =   735
                        Left            =   45
                        TabIndex        =   38
                        Top             =   630
                        Width           =   4110
                        Begin VB.TextBox txtCasualtiesAffected 
                           DataField       =   "Casualties Affected"
                           Height          =   285
                           Left            =   2790
                           TabIndex        =   42
                           Top             =   315
                           Width           =   360
                        End
                        Begin VB.TextBox txtCasualtiesInjured 
                           DataField       =   "Casualties  Injured"
                           Height          =   285
                           Left            =   1755
                           TabIndex        =   41
                           Top             =   315
                           Width           =   360
                        End
                        Begin VB.TextBox txtCasualtiesDead 
                           DataField       =   "Casualties Dead"
                           Height          =   285
                           Left            =   3600
                           TabIndex        =   40
                           Top             =   315
                           Width           =   360
                        End
                        Begin VB.CheckBox chkUnknown 
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Unknown"
                           Height          =   285
                           Left            =   90
                           TabIndex        =   39
                           Top             =   315
                           Width           =   1005
                        End
                        Begin VB.Label lblAffected 
                           AutoSize        =   -1  'True
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Affected:"
                           Height          =   195
                           Left            =   2115
                           TabIndex        =   45
                           Top             =   405
                           Width           =   645
                        End
                        Begin VB.Label lblDead 
                           AutoSize        =   -1  'True
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Dead:"
                           Height          =   195
                           Left            =   3150
                           TabIndex        =   44
                           Top             =   405
                           Width           =   435
                        End
                        Begin VB.Label lblInjured 
                           AutoSize        =   -1  'True
                           BackColor       =   &H00FFFFFF&
                           Caption         =   "Injured:"
                           Height          =   195
                           Left            =   1215
                           TabIndex        =   43
                           Top             =   360
                           Width           =   525
                        End
                     End
                     Begin VB.OptionButton OptCasualties 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "Yes"
                        Height          =   330
                        Index           =   0
                        Left            =   90
                        TabIndex        =   37
                        Top             =   270
                        Width           =   1050
                     End
                     Begin VB.OptionButton OptCasualties 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "No"
                        Height          =   330
                        Index           =   1
                        Left            =   1395
                        TabIndex        =   36
                        Top             =   270
                        Width           =   1050
                     End
                     Begin VB.OptionButton OptCasualties 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "Unknown"
                        Height          =   330
                        Index           =   2
                        Left            =   2745
                        TabIndex        =   35
                        Top             =   270
                        Value           =   -1  'True
                        Width           =   1050
                     End
                  End
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elTime 
            Height          =   4365
            Left            =   9165
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   315
            Width           =   8520
            _cx             =   15028
            _cy             =   7699
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
            Begin VB.Frame FraTimeOf 
               Caption         =   "Time Of Incident:"
               Height          =   2805
               Left            =   3015
               TabIndex        =   145
               Top             =   135
               Width           =   5370
               Begin VB.Frame FraExactTime 
                  Caption         =   "Exact Time:"
                  Height          =   1185
                  Left            =   180
                  TabIndex        =   146
                  Top             =   360
                  Width           =   1140
                  Begin VB.TextBox txtMinutes 
                     Height          =   285
                     Left            =   585
                     MaxLength       =   2
                     TabIndex        =   148
                     Text            =   "00"
                     Top             =   540
                     Width           =   330
                  End
                  Begin VB.TextBox txtHour 
                     Height          =   285
                     Left            =   135
                     MaxLength       =   2
                     TabIndex        =   147
                     Text            =   "12"
                     Top             =   540
                     Width           =   285
                  End
                  Begin VB.Label lblSeparator 
                     AutoSize        =   -1  'True
                     Caption         =   ":"
                     Height          =   195
                     Left            =   495
                     TabIndex        =   151
                     Top             =   540
                     Width           =   45
                  End
                  Begin VB.Label lblEG 
                     AutoSize        =   -1  'True
                     Caption         =   "e.g. 14:37"
                     Height          =   195
                     Left            =   135
                     TabIndex        =   150
                     Top             =   900
                     Width           =   720
                  End
                  Begin VB.Label lblHourMinutes 
                     AutoSize        =   -1  'True
                     Caption         =   "Hour:Minutes"
                     Height          =   195
                     Left            =   90
                     TabIndex        =   149
                     Top             =   270
                     Width           =   945
                  End
               End
            End
            Begin VB.Frame FraDateOf 
               Caption         =   "Date Of Incident:"
               Height          =   2805
               Left            =   90
               TabIndex        =   143
               Top             =   135
               Width           =   2895
               Begin MSComCtl2.MonthView MVIncident 
                  Height          =   2370
                  Left            =   90
                  TabIndex        =   144
                  Top             =   270
                  Width           =   2700
                  _ExtentX        =   4763
                  _ExtentY        =   4180
                  _Version        =   393216
                  ForeColor       =   -2147483630
                  BackColor       =   -2147483633
                  Appearance      =   1
                  StartOfWeek     =   80084994
                  CurrentDate     =   39253
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elStep 
            Height          =   4365
            Index           =   0
            Left            =   -9135
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   315
            Width           =   8520
            _cx             =   15028
            _cy             =   7699
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
            Begin VB.ComboBox ComInformationValidity 
               Height          =   315
               Left            =   4920
               Sorted          =   -1  'True
               TabIndex        =   162
               Top             =   1110
               Visible         =   0   'False
               Width           =   3615
            End
            Begin VB.ComboBox ComSourceReliability 
               Height          =   315
               Left            =   4920
               Sorted          =   -1  'True
               TabIndex        =   160
               Top             =   600
               Visible         =   0   'False
               Width           =   3615
            End
            Begin VB.ComboBox txtEnteredBy 
               Height          =   315
               Left            =   990
               TabIndex        =   158
               Top             =   120
               Width           =   2715
            End
            Begin VB.ComboBox ComSource 
               Height          =   315
               Left            =   4920
               Sorted          =   -1  'True
               TabIndex        =   157
               Top             =   150
               Width           =   3615
            End
            Begin MSComCtl2.MonthView mvDateEntered 
               Height          =   2370
               Left            =   120
               TabIndex        =   142
               Top             =   960
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   4180
               _Version        =   393216
               ForeColor       =   -2147483630
               BackColor       =   -2147483633
               Appearance      =   1
               StartOfWeek     =   80084994
               CurrentDate     =   39326
            End
            Begin VB.TextBox txtEntryDate 
               Height          =   330
               Left            =   990
               TabIndex        =   20
               Text            =   "entryDate"
               Top             =   495
               Visible         =   0   'False
               Width           =   2400
            End
            Begin VB.Label lblInformationValidity 
               Caption         =   "Information Validity:"
               Height          =   465
               Left            =   4050
               TabIndex        =   161
               Top             =   1020
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label lblSourceReliability 
               Caption         =   "Source Reliability:"
               Height          =   465
               Left            =   4110
               TabIndex        =   159
               Top             =   510
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblSource 
               AutoSize        =   -1  'True
               Caption         =   "Source:"
               Height          =   195
               Left            =   4110
               TabIndex        =   156
               Top             =   180
               Width           =   555
            End
            Begin VB.Label lblEntryDate 
               Caption         =   "Entry Date:"
               Height          =   240
               Left            =   45
               TabIndex        =   19
               Top             =   585
               Width           =   960
            End
            Begin VB.Label lblEnteredby 
               Caption         =   "Entered by:"
               Height          =   285
               Left            =   45
               TabIndex        =   18
               Top             =   180
               Width           =   870
            End
         End
         Begin C1SizerLibCtl.C1Elastic elIntro 
            Height          =   4365
            Left            =   -9435
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   315
            Width           =   8520
            _cx             =   15028
            _cy             =   7699
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   4365
               Left            =   0
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   0
               Width           =   8520
               _cx             =   15028
               _cy             =   7699
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
               GridRows        =   5
               GridCols        =   1
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   0
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmAddIncident.frx":13B35
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic C1Logo 
                  Height          =   1530
                  Left            =   0
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   8520
                  _cx             =   15028
                  _cy             =   2699
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
                  BackColor       =   16777215
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
                  FrameWidth      =   0
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   ""
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin VB.Image Image1 
                     Height          =   1185
                     Left            =   0
                     Picture         =   "frmAddIncident.frx":13B96
                     Stretch         =   -1  'True
                     Top             =   120
                     Width           =   4800
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1EOperationsActivities 
                  Height          =   2040
                  Left            =   0
                  TabIndex        =   23
                  TabStop         =   0   'False
                  Top             =   1530
                  Width           =   8520
                  _cx             =   15028
                  _cy             =   3598
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial Black"
                     Size            =   26.25
                     Charset         =   0
                     Weight          =   900
                     Underline       =   0   'False
                     Italic          =   -1  'True
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
                  Caption         =   "Operations Activity Security Information System"
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
                  FrameWidth      =   0
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   ""
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
               End
               Begin VB.Label lblClientVersion 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   0
                  TabIndex        =   26
                  Top             =   3570
                  Width           =   8520
               End
               Begin VB.Label lblWwwHumanitariansecurity 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "www.immap.org"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   0
                  TabIndex        =   25
                  Top             =   4080
                  Width           =   8520
               End
               Begin VB.Label lblDevelopedBu 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Developed by IMMAP Inc. and IMMAP Europe"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   0
                  TabIndex        =   24
                  Top             =   3840
                  Width           =   8520
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elStep 
            Height          =   4365
            Index           =   1
            Left            =   15
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   315
            Width           =   8520
            _cx             =   15028
            _cy             =   7699
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
            Begin VB.CheckBox chkUseDistrict 
               Caption         =   "Use District Level Only"
               Height          =   375
               Left            =   3900
               TabIndex        =   164
               Top             =   780
               Width           =   1575
            End
            Begin VB.TextBox txtLocationDescription 
               Height          =   1785
               Left            =   135
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   62
               Top             =   2070
               Width           =   5205
            End
            Begin VB.CommandButton cmdCheckIn 
               Height          =   510
               Left            =   3870
               Picture         =   "frmAddIncident.frx":1A030
               Style           =   1  'Graphical
               TabIndex        =   59
               ToolTipText     =   "Check In Map"
               Top             =   1260
               Width           =   600
            End
            Begin VB.CommandButton cmdGetCoordinate 
               Height          =   510
               Left            =   3870
               Picture         =   "frmAddIncident.frx":20882
               Style           =   1  'Graphical
               TabIndex        =   58
               ToolTipText     =   "Get Coordinate from Map"
               Top             =   180
               Width           =   600
            End
            Begin VB.Frame FraAutoLocator 
               Caption         =   "Auto Locator"
               Height          =   1080
               Left            =   5715
               TabIndex        =   54
               Top             =   90
               Width           =   2760
               Begin VB.Label lblProvince_ 
                  AutoSize        =   -1  'True
                  Caption         =   "Province:____________________"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   57
                  Top             =   225
                  Width           =   2475
               End
               Begin VB.Label lblDistrict_ 
                  AutoSize        =   -1  'True
                  Caption         =   "District:______________________"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   56
                  Top             =   450
                  Width           =   2505
               End
               Begin VB.Label lblNearestTown 
                  AutoSize        =   -1  'True
                  Caption         =   "Nearest Town:________________"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   55
                  Top             =   675
                  Width           =   2490
               End
            End
            Begin VB.Frame FraCoordinate 
               Caption         =   "Location"
               Height          =   1680
               Left            =   135
               TabIndex        =   5
               Top             =   90
               Width           =   3660
               Begin VB.TextBox txtMGRS 
                  Height          =   285
                  Left            =   720
                  TabIndex        =   11
                  Top             =   1305
                  Width           =   2715
               End
               Begin VB.Frame FraXY 
                  Caption         =   "X:Y"
                  Height          =   1005
                  Left            =   135
                  TabIndex        =   6
                  Top             =   270
                  Width           =   3300
                  Begin VB.TextBox txtX 
                     Height          =   285
                     Left            =   1260
                     TabIndex        =   8
                     Top             =   270
                     Width           =   1905
                  End
                  Begin VB.TextBox txtY 
                     Height          =   285
                     Left            =   1260
                     TabIndex        =   7
                     Top             =   630
                     Width           =   1905
                  End
                  Begin VB.Label lblX 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "X"
                     Height          =   195
                     Left            =   1035
                     TabIndex        =   10
                     Top             =   270
                     Width           =   105
                  End
                  Begin VB.Label lblY 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Y"
                     Height          =   195
                     Left            =   1035
                     TabIndex        =   9
                     Top             =   630
                     Width           =   105
                  End
               End
               Begin VB.Label lblMGRS 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "MGRS:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   12
                  Top             =   1350
                  Width           =   555
               End
            End
            Begin VB.Label lblNearestOp 
               AutoSize        =   -1  'True
               Caption         =   "Nearest Op Area:________________"
               Height          =   195
               Left            =   5760
               TabIndex        =   166
               Top             =   1560
               Visible         =   0   'False
               Width           =   2490
            End
            Begin VB.Label lblNearestOffice 
               AutoSize        =   -1  'True
               Caption         =   "Nearest Office:________________"
               Height          =   195
               Left            =   5760
               TabIndex        =   165
               Top             =   1320
               Visible         =   0   'False
               Width           =   2490
            End
            Begin VB.Label lblLocationDescription 
               Caption         =   "Location Description"
               Height          =   285
               Left            =   90
               TabIndex        =   63
               Top             =   1800
               Width           =   2760
            End
            Begin VB.Label lblGetCoordinates 
               Caption         =   "Get Position From Map"
               Height          =   510
               Left            =   4590
               TabIndex        =   61
               Top             =   180
               Width           =   1125
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblCheckYour 
               Caption         =   "Check Your position"
               Height          =   510
               Left            =   4590
               TabIndex        =   60
               Top             =   1260
               Width           =   1095
               WordWrap        =   -1  'True
            End
         End
      End
   End
End
Attribute VB_Name = "frmAddIncident"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event GetCoordinatesFromMap(sAdmVal1 As String, sAdmVal2 As String, sAdmloc As String)
Public Event CheckCoordinates(sAdmVal1 As String, sAdmVal2 As String, sAdmloc As String)
Public Event SubmitIncident(sGUID As String) '(RSVictims As ADODB.Recordset)
Public Event CancelIncident()
Public Event ConvPTtoMGRS(x As Double, y As Double)
Public Event ConvMGRStoPT(sMGRS As String, x As Double, y As Double)

Private m_sAdmVal1 As String
Private m_sAdmVal2 As String
Private m_sAdmLoc As String

'Private RSVictims As ADODB.Recordset
Private m_Conn As ADODB.Connection
Private m_bUnknown As Boolean
Private m_intIncidentCharacter As Integer

Private ToolT_txtCasualtiesAffected As New clsTooltips
Private ToolT_txtCasualtiesDead As New clsTooltips
Private ToolT_txtCasualtiesInjured As New clsTooltips
Private ToolT_txtEnteredBy As New clsTooltips
Private ToolT_txtEntryDate As New clsTooltips
Private ToolT_txtHour As New clsTooltips
Private ToolT_txtIncidentDescription As New clsTooltips
Private ToolT_txtLocationDescription As New clsTooltips
Private ToolT_txtMGRS As New clsTooltips
Private ToolT_txtMinutes As New clsTooltips
Private ToolT_txtX As New clsTooltips
Private ToolT_txtY As New clsTooltips
Private ToolT_cmdAdd As New clsTooltips
Private ToolT_cmdBack As New clsTooltips
Private ToolT_cmdCancel As New clsTooltips
Private ToolT_cmdCheckIn As New clsTooltips
Private ToolT_cmdGetCoordinate As New clsTooltips
Private ToolT_cmdSubmit As New clsTooltips
Private ToolT_cmdWizValUpdate As New clsTooltips
Private ToolT_ComIncTarget As New clsTooltips
Private ToolT_ComIncType As New clsTooltips
Private ToolT_MVIncident As New clsTooltips

Private WithEvents m_cFileCopy As clsFileCopy
Attribute m_cFileCopy.VB_VarHelpID = -1
Private m_sAttachments() As String

Dim acSource As New CAutoComplete
Dim acEntered As New CAutoComplete
Private m_sGUID As String

Public Property Get IncAttachments() As Variant
    IncAttachments = m_sAttachments
End Property

Private Sub CreateToolTips()
    ToolT_txtCasualtiesAffected.CreateBalloon txtCasualtiesAffected, "Number of casualties affected." _
    & vbCrLf & "Indicates how many people are being affected by an incident." _
    & vbCrLf & "e.g. if a powerstation is being attached.", "Incident Detail Section", 1
    
    ToolT_txtCasualtiesDead.CreateBalloon txtCasualtiesDead, "Number of casualties dead.", "Incident Detail Section", 1
    
    ToolT_txtCasualtiesInjured.CreateBalloon txtCasualtiesInjured, "Number of casualties injured.", "Incident Detail Section", 1
    
    ToolT_txtEnteredBy.CreateBalloon txtEnteredBy, "Who the incident report is entered by.", "Report Details", 1
    
    ToolT_txtEntryDate.CreateBalloon txtEntryDate, "The date the report is entered", "Report Details", 1
    
    ToolT_txtHour.CreateBalloon txtHour, "Indicate the hour incident happend.", "Incident Detail Section", 1
    
    ToolT_txtIncidentDescription.CreateBalloon txtIncidentDescription, "Describe the incident as detailed as possible", "Incident Detail Section", 1
    
    ToolT_txtLocationDescription.CreateBalloon txtLocationDescription, "Describe the location as detailed as possible." _
    & vbCrLf & "Include characteristics and features that might have had" _
    & vbCrLf & "impct on the incident outcome.", "Incident Location Section", 1
    
    ToolT_txtMGRS.CreateBalloon txtMGRS, "MGRS - Military Grid Reference System" _
    & vbCrLf & "Indicates the incident position in MGRS" _
    & vbCrLf & "This is automatically filled in once Latitude and Longitude" _
    & vbCrLf & "is being filled in", "Incident Location Section", 1
    
    ToolT_txtMinutes.CreateBalloon txtMinutes, "Indicate the approximate minutes when the Incident happend", "Incident Detail Section", 1
    
    ToolT_txtX.CreateBalloon txtX, "Fill in the X/Longitude/Easting." _
    & vbCrLf & "Should be in Decimal Degrees Format:" _
    & vbCrLf & "e.g xx.xxxxxx", "Incident Location Section", 1
    
    ToolT_txtY.CreateBalloon txtY, "Fill in the Y/Latitude/Northing." _
    & vbCrLf & "Should be in Decimal Degrees Format:" _
    & vbCrLf & "e.g xx.xxxxxx", "Incident Location Section", 1
    
    ToolT_cmdAdd.CreateBalloon cmdAdd, "Add the Item", "Navigation", 1
    
    ToolT_cmdBack.CreateBalloon cmdBack, "Navigate Back In the Wizard", "Navigation", 1
    
    ToolT_cmdCancel.CreateBalloon cmdCancel, "Cancel item", "Navigation", 1
    
    ToolT_cmdCheckIn.CreateBalloon cmdCheckIn, "Check", "Navigation", 1
    
    ToolT_cmdGetCoordinate.CreateBalloon cmdGetCoordinate, "Click in the map to get coordinate", "Incident Location Section", 1
    
    ToolT_cmdSubmit.CreateBalloon cmdSubmit, "Submit your reported incident", "Navigation", 1
    
    'ToolT_cmdWizValUpdate.CreateBalloon cmdWizValUpdate, "wtwtwt", "wtwtwttwwt", 1
    
    ToolT_ComIncTarget.CreateBalloon ComIncTarget, "Choose the Incident target that best apply", "Incident Detail Section", 1
    
    ToolT_ComIncType.CreateBalloon ComIncType, "Choose the Incident type that best apply", "Incident Detail Section", 1
End Sub


Public Sub SetCoordinateValue(x As String, xLabel As String, y As String, yLabel As String, sCoordSysName As String, sMGRS As String)
    txtX.Text = x
    lblX.caption = xLabel
    txtY.Text = y
    lblY.caption = yLabel
    txtMGRS.Text = sMGRS
    FraXY.caption = sCoordSysName
End Sub

Public Sub SetNearestValues(dNearestOffice As Double, dNearestOP As Double)
    'lblNearestOffice = "Nearest Office: " & Round(dNearestOffice / 1000, 2) & " km"
    'lblNearestOp = "Nearest OP Area: " & Round(dNearestOP / 1000, 2) & " km"
End Sub

Public Sub SetAdmValues(sAdmVal1 As String, _
                        sAdmVal2 As String, _
                        sAdmloc As String)

        With g_RSAppSettings
      
6110        .Requery
6112        SafeMoveFirst g_RSAppSettings
6114        .Find "SettingName = 'AdmProvSec'"

6115        If Not .Fields.Item("SettingValue7").Value = vbNullString Then lblProvince_.caption = .Fields.Item("SettingValue7").Value
        
            If sAdmVal1 = "" Then
                If Not .Fields.Item("SettingValue7").Value = vbNullString Then
                    lblProvince_.caption = .Fields.Item("SettingValue7").Value & ": N/A"
                Else
                    lblProvince_.caption = "Province: N/A"
                End If

            Else

                If Not .Fields.Item("SettingValue7").Value = vbNullString Then
                    lblProvince_.caption = .Fields.Item("SettingValue7").Value & ": " & sAdmVal1
                Else
                    lblProvince_.caption = "Province: " & sAdmVal1
                End If
            End If

            SafeMoveFirst g_RSAppSettings
            .Find "SettingName = 'AdmDistSec'"
            
            If sAdmVal2 = "" Then
                If Not .Fields.Item("SettingValue7").Value = vbNullString Then
                    lblProvince_.caption = .Fields.Item("SettingValue7").Value & ": N/A"
                Else
                    lblDistrict_.caption = "District: N/A"
                End If

            Else

                If Not .Fields.Item("SettingValue7").Value = vbNullString Then
                    lblDistrict_.caption = .Fields.Item("SettingValue7").Value & ": " & sAdmVal2
                Else
                    lblDistrict_.caption = "District: " & sAdmVal2
                End If
            End If

            SafeMoveFirst g_RSAppSettings
            .Find "SettingName = 'AdmLocalSec'"

            If Not .Fields.Item("SettingValue7").Value = vbNullString Then lblNearestTown.caption = .Fields.Item("SettingValue7").Value

            If sAdmloc = "" Then
                If Not .Fields.Item("SettingValue7").Value = vbNullString Then
                    lblNearestTown.caption = .Fields.Item("SettingValue7").Value & ": N/A"
                Else
                    lblNearestTown.caption = "Nearest Town: N/A"
                End If

            Else

                If Not .Fields.Item("SettingValue7").Value = vbNullString Then
                    lblNearestTown.caption = .Fields.Item("SettingValue7").Value & ": " & sAdmloc
                Else
                    lblNearestTown.caption = "Nearest Town: " & sAdmloc
                End If
            End If

        End With

End Sub

Public Sub ResetValues()
    ComIncTarget.ListIndex = 0
    ComIncType.ListIndex = 0
    txtIncidentDescription.Text = ""
    txtLocationDescription.Text = ""
    txtX.Text = ""
    txtY.Text = ""
    txtMGRS.Text = ""
 
    OptCasualties(2).Value = True
    OptIncViolence(2).Value = True
    chkUnknown.Value = vbChecked
    txtCasualtiesAffected.Text = ""
    txtCasualtiesDead.Text = ""
    txtCasualtiesInjured.Text = ""
    lstAttachments.Clear
    ReDim m_sAttachments(0)

'    If RSVictims.State = adStateOpen Then RSVictims.Close
'    Set RSVictims = New ADODB.Recordset
'    RSVictims.Fields.Append "#", adInteger
'    RSVictims.Fields.Append "Occupation", adVarChar, 80
'    RSVictims.Fields.Append "Under18", adVarChar, 10
'    RSVictims.Fields.Append "Sex", adVarChar, 10
'    RSVictims.Fields.Append "Condition", adVarChar, 10
'    RSVictims.Fields.Append "Ethnicity", adVarChar, 40
'    RSVictims.Fields.Append "Quantity", adBigInt
'    RSVictims.Open
'    dxDBGrid.Columns.DestroyColumns
'    Set Me.dxDBGrid.DataSource = RSVictims
'    dxDBGrid.Columns.RetrieveFields
    
End Sub

Private Sub c1TabAddPt_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If NewTab = c1TabAddPt.NumTabs - 1 Then
    
        lblWizValLabel(0).caption = "Entered By: " & txtEnteredBy.Text
        lblWizValLabel(1).caption = "Entry Date: " & mvDateEntered.Value
        lblWizValLabel(2).caption = "X: " & txtX.Text
        lblWizValLabel(3).caption = "Y: " & txtY.Text
        lblWizValLabel(4).caption = "MGRS: " & txtMGRS.Text
        lblWizValLabel(5).caption = "Location Desc: " & txtLocationDescription.Text
        lblWizValLabel(6).caption = "Date Of Incident: " & MVIncident.Value
        lblWizValLabel(7).caption = "Time of Incident: " & txtHour.Text & ":" & txtMinutes.Text
        
        If OptIncViolence(0).Value Then
            lblWizValLabel(8).caption = "Incident character: Violent"
        ElseIf OptIncViolence(1).Value Then
            lblWizValLabel(8).caption = "Incident character: Non Violent"
        Else
            lblWizValLabel(8).caption = "Incident character: Unknown"
        End If
        
        
        lblWizValLabel(9).caption = "Type: " & ComIncType.List(ComIncType.ListIndex)
        lblWizValLabel(10).caption = "Target: " & ComIncTarget.List(ComIncTarget.ListIndex)
        lblWizValLabel(11).caption = "# Injured: " & txtCasualtiesInjured.Text
        lblWizValLabel(12).caption = "# Dead: " & txtCasualtiesDead.Text
        lblWizValLabel(13).caption = "Incident Desc: " & txtIncidentDescription.Text
        lblWizValLabel(14).caption = "Verify when published: " & IIf(chkSendMe.Value = vbChecked, "Yes", "No")
        lblWizValLabel(15).caption = "Allows others to contact me: " & IIf(chkAllowOther.Value = vbChecked, "Yes", "No")
        lblWizValLabel(16).caption = "Attachments"
    
    End If
    
    If OldTab = 5 And NewTab = 6 And c1TabAddPt.TabVisible(6) = False Then
        NewTab = 7
    ElseIf OldTab = 7 And NewTab = 6 And c1TabAddPt.TabVisible(6) = False Then
        NewTab = 5
    End If
    
    If NewTab = c1TabAddPt.NumTabs - 1 Then
        cmdForward.Enabled = False
        cmdSubmit.Enabled = True
    Else
        cmdForward.Enabled = True
        cmdSubmit.Enabled = False
    End If
    
    
End Sub



Private Sub chkUnknown_Click()
    lblAffected.Enabled = IIf(chkUnknown.Value = vbUnchecked, True, False)
    lblDead.Enabled = IIf(chkUnknown.Value = vbUnchecked, True, False)
    lblInjured.Enabled = IIf(chkUnknown.Value = vbUnchecked, True, False)
    txtCasualtiesAffected.Enabled = IIf(chkUnknown.Value = vbUnchecked, True, False)
    txtCasualtiesDead.Enabled = IIf(chkUnknown.Value = vbUnchecked, True, False)
    txtCasualtiesInjured.Enabled = IIf(chkUnknown.Value = vbUnchecked, True, False)
    
    If chkUnknown.Value = vbChecked Then
    
        txtCasualtiesAffected = ""
        txtCasualtiesDead = ""
        txtCasualtiesInjured = ""
        
    Else
    
            txtCasualtiesAffected = 0
        txtCasualtiesDead = 0
        txtCasualtiesInjured = 0
        
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim c As New cCommonDialog

    With c
                    
        'Set filter to all files
        .Filter = "All files|*.*"
        'Show the open file window
        .ShowOpen
        'Put the file bath in the source text
        lstAttachments.AddItem .Filename
        'TxtSrc.Text = CmnDlg.Filename
    End With

End Sub

Private Sub cmdAddAttachments_Click()
Dim sAddress() As String
Dim i As Integer

    sAddress = Split(g_sAppServerPath, ".")

    g_ssubdomain = Replace(sAddress(0), "http://", "")

    With frmAttachmentUploader.OASISAttUploader
        .hostname = "www.oasiswebservice.org"
        .ChunkSize = 2048
        .IdleTimeOutMs = 60000
        .Port = 80
        .RecordGUID = m_sGUID
        .ServerPath = "/upl/"
        .UseBatchMode = True
        .UserGroup = g_sRemoteTablePrefix
        .TableName = "oincidents"
        .subdomainname = g_ssubdomain
        frmAttachmentUploader.Show vbModal, Me
    End With
End Sub

'Private Sub cmdAddVictim_Click()
'
'    RSVictims.AddNew
'    RSVictims.Fields(0).Value = RSVictims.RecordCount
'    RSVictims.Fields(1).Value = Me.cmbVictim(0)
'    RSVictims.Fields(2).Value = Me.cmbVictim(1)
'    RSVictims.Fields(3).Value = Me.cmbVictim(2)
'    RSVictims.Fields(4).Value = Me.cmbVictim(3)
'    RSVictims.Fields(5).Value = Me.cmbVictim(4)
'    RSVictims.Fields(6).Value = Me.dxQuantity
'
'    'frmIncidentVictimAnalysis.Show
'
'End Sub

Private Sub cmdBack_Click()

    If c1TabAddPt.CurrTab > 0 Then
        If c1TabAddPt.CurrTab = 6 Then
            c1TabAddPt.CurrTab = c1TabAddPt.CurrTab - 2
        Else
            c1TabAddPt.CurrTab = c1TabAddPt.CurrTab - 1
        End If
    End If
    cmdForward.Enabled = True
    cmdSubmit.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    cmdForward.Enabled = True
    cmdSubmit.Enabled = False
    RaiseEvent CancelIncident
End Sub

Private Sub cmdCheckIn_Click()
    RaiseEvent CheckCoordinates(m_sAdmVal1, m_sAdmVal2, m_sAdmLoc)
End Sub

Private Sub cmdForward_Click()

    DebugPrint Me.Controls.Count

    If c1TabAddPt.CurrTab < c1TabAddPt.NumTabs - 1 Then
        If c1TabAddPt.CurrTab = 4 Then
            c1TabAddPt.CurrTab = c1TabAddPt.CurrTab + 2
        Else
            c1TabAddPt.CurrTab = c1TabAddPt.CurrTab + 1
        End If
    End If
    
    If c1TabAddPt.CurrTab = c1TabAddPt.NumTabs - 1 Then
        cmdForward.Enabled = False
        cmdSubmit.Enabled = True
    Else
        cmdForward.Enabled = True
        cmdSubmit.Enabled = False
    End If
    
    lblWizValLabel(0).caption = "Entered By: " & txtEnteredBy.Text
    lblWizValLabel(1).caption = "Entry Date: " & mvDateEntered.Value
    lblWizValLabel(2).caption = "Lat: " & txtX.Text 'These are twisted!!!
    lblWizValLabel(3).caption = "Long: " & txtY.Text
    lblWizValLabel(4).caption = "MGRS: " & txtMGRS.Text
    lblWizValLabel(5).caption = "Location Desc:" & txtLocationDescription.Text
    lblWizValLabel(6).caption = "Date Of Incident:" & MVIncident.Value
    lblWizValLabel(7).caption = "Time of Incident:" & txtHour.Text & ":" & txtMinutes.Text
        
    If OptIncViolence(0).Value Then
        lblWizValLabel(8).caption = "Incident character: Violent"
    ElseIf OptIncViolence(1).Value Then
        lblWizValLabel(8).caption = "Incident character: Non Violent"
    Else
        lblWizValLabel(8).caption = "Incident character: Unknown"
    End If
        
    lblWizValLabel(9).caption = "Type: " & ComIncType.List(ComIncType.ListIndex)
    lblWizValLabel(10).caption = "Target: " & ComIncTarget.List(ComIncTarget.ListIndex)
    lblWizValLabel(11).caption = "# Injured: " & txtCasualtiesInjured.Text
    lblWizValLabel(12).caption = "# Dead: " & txtCasualtiesDead.Text
    lblWizValLabel(13).caption = "Incident Desc: " & txtIncidentDescription.Text
    lblWizValLabel(14).caption = "Verify when published: " & IIf(chkSendMe.Value = vbChecked, "Yes", "No")
    lblWizValLabel(15).caption = "Allows others to contact me: " & IIf(chkAllowOther.Value = vbChecked, "Yes", "No")
    lblWizValLabel(16).caption = "Attachments"

End Sub

Private Sub cmdGetCoordinate_Click()
    RaiseEvent GetCoordinatesFromMap(m_sAdmVal1, m_sAdmVal2, m_sAdmLoc)
End Sub

Private Sub cmdSubmit_Click()

    If ValidateValues Then
        
        Me.cmdSubmit.Enabled = False
        RaiseEvent SubmitIncident(m_sGUID) '(RSVictims)
        Me.cmdForward.Enabled = True
        
    End If
    
End Sub

Private Sub cmdWizValUpdate_Click(Index As Integer)
    Select Case Index
    
        Case 0, 1
            c1TabAddPt.CurrTab = 1
            If Index = 0 Then
                txtEnteredBy.SetFocus
            Else
                mvDateEntered.SetFocus
            End If
        Case 2, 3, 4, 5
            c1TabAddPt.CurrTab = 2
            If Index = 2 Then
                txtX.SetFocus
            ElseIf Index = 3 Then
                txtY.SetFocus
            ElseIf Index = 4 Then
                txtMGRS.SetFocus
            Else
                txtLocationDescription.SetFocus
            End If
        Case 6, 7
            c1TabAddPt.CurrTab = 3
            If Index = 6 Then
                MVIncident.SetFocus
            Else
                txtHour.SetFocus
            End If
        Case 8, 9, 10, 11, 12, 13
            c1TabAddPt.CurrTab = 4
        Case 14, 15, 16
            c1TabAddPt.CurrTab = 5
    End Select
End Sub


Private Sub Form_Load()
    
    Dim sString(5) As String
    Dim sIniFile As String
    Dim i As Integer
    
    If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
    
'    Set RSVictims = New ADODB.Recordset
'    RSVictims.Fields.Append "#", adInteger
'    RSVictims.Fields.Append "Occupation", adVarChar, 80
'    RSVictims.Fields.Append "Under18", adVarChar, 10
'    RSVictims.Fields.Append "Sex", adVarChar, 10
'    RSVictims.Fields.Append "Condition", adVarChar, 10
'    RSVictims.Fields.Append "Ethnicity", adVarChar, 40
'    RSVictims.Fields.Append "Quantity", adBigInt
'    RSVictims.Open
'
'    sIniFile = g_sAppPath & "\data\user\incidentvictims.ini"
'
'    If CheckIfFileExists(sIniFile) Then
'
'        sString(0) = ReadIniValue(sIniFile, "Default", "Occupation")
'        sString(1) = ReadIniValue(sIniFile, "Default", "Under18")
'        sString(2) = ReadIniValue(sIniFile, "Default", "Sex")
'        sString(3) = ReadIniValue(sIniFile, "Default", "Condition")
'        sString(4) = ReadIniValue(sIniFile, "Default", "Ethnicity")
'
'        i = 0
'
'        Do Until i = 5
'
'            Do
'                If Not InStr(sString(i), ":::") = 0 Then
'                    cmbVictim(i).AddItem Left(sString(i), InStr(sString(i), ":::") - 1)
'                    sString(i) = Right$(sString(i), Len(sString(i)) - InStr(sString(i), ":::") - 2)
'                Else
'                    cmbVictim(i).AddItem sString(i)
'                    sString(i) = ""
'                End If
'
'            Loop Until InStr(sString(i), ":::") = 0
'
'            cmbVictim(i).Text = cmbVictim(i).List(0)
'            i = i + 1
'        Loop
'
'        dxDBGrid.Columns.DestroyColumns
'        Set Me.dxDBGrid.DataSource = RSVictims
'        dxDBGrid.Columns.RetrieveFields
'    Else
'
'    'HIDE STUFF
'        c1TabAddPt.TabVisible(6) = False
'
'    End If
    Me.cmdForward.Enabled = True
    Me.cmdSubmit.Enabled = False
    Me.txtEntryDate = Format(Now(), "Medium Date")
    MVIncident = Format(Now(), "Medium Date")

    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If

End Sub

'Private Sub Form_Unload(Cancel As Integer)
'
'    If RSVictims.State = adStateOpen Then RSVictims.Close
'
'End Sub



Private Sub mvDateEntered_DateClick(ByVal DateClicked As Date)
    txtEntryDate.Text = mvDateEntered.Value
End Sub

Private Sub OptCasualties_Click(Index As Integer)
    If Index = 0 Then
        FraNumberOf.Enabled = True
        chkUnknown.Value = vbUnchecked
        Call chkUnknown_Click
    Else
        FraNumberOf.Enabled = False
        chkUnknown.Value = vbChecked
        Call chkUnknown_Click
    End If
End Sub

Public Sub Init(oConn As ADODB.Connection, _
                userID As String, _
                Optional sDefaultAdmLev1 As String, _
                Optional sDefaultAdmLev2 As String, _
                Optional sDefaultAdmLoc As String)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        Dim RS As New ADODB.Recordset
     
        m_sGUID = GetGuid
     
100     RS.Open "SELECT * FROM [IncTarget] ORDER BY [name]", oConn, adOpenForwardOnly, adLockReadOnly
    
102     If Not RS.Bof Then SafeMoveFirst RS
    
104     ComIncTarget.Clear
        ComIncTarget.AddItem "-- SELECT A TARGET --"

106     Do While Not RS.EOF

108         If Not RS.Fields.Item("Name").Value = vbNull Then ComIncTarget.AddItem RS.Fields.Item("Name").Value
        
110         RS.MoveNext
        Loop
    
112     ComIncTarget.ListIndex = 0
    
114     RS.Close
        Set RS = New ADODB.Recordset
116     RS.Open "SELECT * FROM IncTypeCategory ORDER BY Incident_Type_Name", oConn, adOpenForwardOnly, adLockReadOnly
    
118     If Not RS.Bof Then SafeMoveFirst RS
    
120     ComIncType.Clear
        ComIncType.AddItem "--SELECT AN INCIDENT TYPE--"

122     Do While Not RS.EOF

124         If Not RS.Fields.Item("Incident_Type_Name").Value = vbNull Then ComIncType.AddItem RS.Fields.Item("Incident_Type_Name").Value
126         RS.MoveNext
        Loop
    
        On Error Resume Next
    
128     ComIncType.ListIndex = 0
    
130     txtCasualtiesAffected.Text = "0"
132     txtCasualtiesDead.Text = "0"
134     txtCasualtiesInjured.Text = "0"
136     ' txtNumberOfSources.Text = ""
138     txtIncidentDescription.Text = ""
        OptCasualties(2).Value = True
        OptIncViolence(2).Value = True
140     RS.Close
142     Set RS = Nothing
        chkUnknown.Value = vbChecked
        
        txtEnteredBy.Text = userID
        mvDateEntered.Value = date
        txtEntryDate.Text = mvDateEntered.Value
        c1TabAddPt.CurrTab = 0
        ResetValues
        
        ComSource.Clear
        txtEnteredBy.Clear
        
        Set acSource.LinkedComboBox = ComSource
        acSource.AutoAdd = True
                
        Set acEntered.LinkedComboBox = txtEnteredBy
        acEntered.AutoAdd = True
        
        CreateToolTips
        
        If OutlookExists Then
            chkCreateEmail.Visible = True
        Else
            chkCreateEmail.Visible = False
        End If
        
        On Error Resume Next
                
        With ComInformationValidity
            .Clear
            .AddItem "Suggested by several independent sources"
            .AddItem "Very likely"
            .AddItem "Likely"
            .AddItem "Not likely"
            .AddItem "Probably wrong information"
            .AddItem "We do not know"
        End With
        
        With ComSourceReliability
            .Clear
            .AddItem "Knowledgeable with direct access to information"
            .AddItem "Knowledgeable but no direct access to information"
            .AddItem "Usually reliable"
            .AddItem "is not usually reliable"
            .AddItem "is not reliable"
            .AddItem "We do not know"
        End With
        
        Set RS = New ADODB.Recordset
        RS.Open "SELECT DISTINCT [Source] FROM oincidents_FEA ORDER BY [source]", oConn

        If Not Err.number > 0 Then

            With RS

                If Not .EOF And Not .Bof Then
                    
                    If SafeMoveFirst(RS) Then
                
                        Do While Not .EOF
                            ComSource.AddItem .Fields.Item("Source").Value
                            .MoveNext
                        Loop
                        
                    End If
                    
                End If

            End With

        End If
        
        Err.Clear
        
        Set RS = New ADODB.Recordset
        RS.Open "SELECT DISTINCT [NAME] FROM oincidents_FEA order by [NAME]", oConn

        If Not Err.number > 0 Then

            If Not RS.State = adStateClosed Then

                With RS

                    If Not .EOF And Not .Bof Then
                        SafeMoveFirst RS

                        Do While Not .EOF
                            txtEnteredBy.AddItem .Fields.Item("NAME").Value
                            .MoveNext
                        Loop

                    End If

                End With

            End If
        
        End If
        
        With g_RSAppSettings
    
6110        .Requery
6112        SafeMoveFirst g_RSAppSettings
6114        .Find "SettingName = 'AdmProvSec'"

6115        If Not .Fields.Item("SettingValue7").Value = vbNullString Then lblProvince_.caption = .Fields.Item("SettingValue7").Value
            SafeMoveFirst g_RSAppSettings
            .Find "SettingName = 'AdmDistSec'"
            If Not .Fields.Item("SettingValue7").Value = vbNullString Then lblDistrict_.caption = .Fields.Item("SettingValue7").Value
            SafeMoveFirst g_RSAppSettings
            .Find "SettingName = 'AdmLocalSec'"
            If Not .Fields.Item("SettingValue7").Value = vbNullString Then lblNearestTown.caption = .Fields.Item("SettingValue7").Value

        End With
        
        '<EhFooter>
        Exit Sub

Init_Err:
        Err.Raise vbObjectError + 100, "OasisReportWizard.IncidentDescription.Init", "IncidentDescription component failure"
        '</EhFooter>
End Sub
Public Function ValidateValues() As Boolean
        '<EhHeader>
        On Error GoTo ValidateValues_Err
        
        ValidateValues = True
        
        If ComIncType.ListIndex = 0 Then
            c1TabAddPt.CurrTab = 4
            ComIncType.SetFocus
            ValidateValues = False
            Err.Raise 101, "Validate Incident Values", "No Incident type selected."
        End If
        
        If ComIncTarget.ListIndex = 0 Then
            c1TabAddPt.CurrTab = 4
            ComIncTarget.SetFocus
            ValidateValues = False
            Err.Raise 101, "Validate Incident Values", "No Incident Target selected."
        End If
        
        If OptCasualties(0).Value = False Then
            ValidateValues = True
        Else
            
        End If
        
       If Not IsNumeric(txtX.Text) Then
            c1TabAddPt.CurrTab = 2
            txtX.SetFocus
            ValidateValues = False
            Err.Raise 101, "Validate Incident Values", "No Valid Incident Coordinate Values Found"
       End If
       
       If Not IsNumeric(txtY.Text) Then
            c1TabAddPt.CurrTab = 2
            txtY.SetFocus
            ValidateValues = False
            Err.Raise 101, "Validate Incident Values", "No Valid Incident Coordinate Values Found"
       End If
       
       If txtMGRS.Text = "" Then
            c1TabAddPt.CurrTab = 2
            txtMGRS.SetFocus
            ValidateValues = False
            Err.Raise 101, "Validate Incident Values", "No Incident MGRS Coordinate Values Found"
       End If
       
       If txtEnteredBy.Text = "" Then
            c1TabAddPt.CurrTab = 1
            txtEnteredBy.SetFocus
            ValidateValues = False
            Err.Raise 101, "Validate Incident Values", "No Entered By Value found"
       End If

       If txtEntryDate.Text = "" Then
            c1TabAddPt.CurrTab = 1
            txtEntryDate.SetFocus
            ValidateValues = False
            Err.Raise 101, "Validate Incident Values", "No Entry Date Value Found"
       End If


        If chkUnknown.Value = vbUnchecked Then
        '</EhHeader>
        
100         If Not txtCasualtiesAffected.Text = "" And Not IsNumeric(txtCasualtiesAffected.Text) Then
102             txtCasualtiesAffected.SelStart = 0
104             txtCasualtiesAffected.SelLength = Len(txtCasualtiesAffected.Text)
106             Err.Raise 101, "Validate Incident Values", "One of the value entered as is not numeric."
            End If
        
108         If Not txtCasualtiesDead.Text = "" And Not IsNumeric(txtCasualtiesDead.Text) Then
110             txtCasualtiesDead.SelStart = 0
112             txtCasualtiesDead.SelLength = Len(txtCasualtiesDead.Text)
114             Err.Raise 101, "Validate Incident Values", "One of the value entered is not numeric."
            End If
        
116         If Not txtCasualtiesInjured.Text = "" And Not IsNumeric(txtCasualtiesInjured.Text) Then
118             txtCasualtiesInjured.SelStart = 0
120             txtCasualtiesInjured.SelLength = Len(txtCasualtiesInjured.Text)
122             Err.Raise 101, "Validate Incident Values", "One of the value entered is not numeric."
            End If
        End If

        Dim j As Integer
        
        If lstAttachments.ListCount > 0 Then
        
            For j = 0 To lstAttachments.ListCount - 1
                DoFileCopy lstAttachments.List(j), g_sAppPath & "\data\user\Attachments\Incidents\" & StripFileName(lstAttachments.List(j)), True
                
                If UBound(m_sAttachments) = 10000 Then
                    ReDim Preserve m_sAttachments(UBound(m_sAttachments) + 1)
                End If
                
                m_sAttachments(UBound(m_sAttachments)) = g_sAppPath & "\data\user\Attachments\Incidents\" & StripFileName(lstAttachments.List(j))
            
                If Not j = lstAttachments.ListCount - 1 Then
                    ReDim Preserve m_sAttachments(UBound(m_sAttachments) + 1)
                End If
            
            Next
            
        End If
            
    
        ValidateValues = True

    
        '<EhFooter>
        Exit Function

ValidateValues_Err:
        MsgBox Err.Description & vbCrLf & "You have to correct the values before you can continue.", vbInformation, "OASIS Incident Report Wizard"
        Err.Clear
        
        ValidateValues = False
        '</EhFooter>
End Function

Private Sub DoFileCopy(TxtSrc As String, TxtDest As String, bVerify As Boolean)
    On Error GoTo CopyErr

    Dim ErrCode As Long
    Dim ErrMsg As String

    Set m_cFileCopy = New clsFileCopy


    'The source file the you want to copy
    m_cFileCopy.SourceFile = TxtSrc
    'The destination file name
    m_cFileCopy.TargetFile = TxtDest

    m_cFileCopy.ByteSize = 4096 '4kb

    'Check for verify
    m_cFileCopy.Verify = bVerify
    
    If Not m_cFileCopy.CopyFile(ErrCode, ErrMsg) Then
        GoTo CopyErr
    End If

    GoTo ex

CopyErr:

    If ErrCode = 0 Then
        ErrCode = Err.number
        ErrMsg = Err.Description
    End If

    MsgBox "Error: " & ErrCode & " - " & ErrMsg, vbCritical, "OASIS Attachment Error"

ex:
    'Close files
    'Close #nSF
    'Close #nDF

    Set m_cFileCopy = Nothing

End Sub

Private Function CheckIfFileExists(Filename As String) As Boolean
    Dim i As Integer
    
    On Local Error Resume Next
    i = Len(Dir$(Filename$))
    If Err Or i = 0 Then
        CheckIfFileExists = False
    Else
        CheckIfFileExists = True
    End If
    On Local Error GoTo 0
End Function


Private Function ReadIniValue(INIPath As String, Key As String, Variable As String) As String
Dim NF As Integer
Dim Temp As String
Dim LcaseTemp As String
Dim ReadyToRead As Boolean
    
AssignVariables:
        NF = FreeFile
        ReadIniValue = ""
        Key = "[" & LCase$(Key) & "]"
        Variable = LCase$(Variable)
    
EnsureFileExists:
    Open INIPath For Binary As NF
    Close NF
    SetAttr INIPath, vbArchive
    
LoadFile:
    Open INIPath For Input As NF
    While Not EOF(NF)
    Line Input #NF, Temp
    LcaseTemp = LCase$(Temp)
    If InStr(LcaseTemp, "[") <> 0 Then ReadyToRead = False
    If LcaseTemp = Key Then ReadyToRead = True
    If InStr(LcaseTemp, "[") = 0 And ReadyToRead = True Then
        If InStr(LcaseTemp, Variable & "=") = 1 Then
            ReadIniValue = Mid$(Temp, 1 + Len(Variable & "="))
            Close NF: Exit Function
            End If
        End If
    Wend
    Close NF
End Function

Private Sub txtMGRS_GotFocus()

    If Not Len(txtMGRS) > 3 And Len(txtY) > 0 And Len(txtX) > 0 Then
        
        If IsNumeric(txtY) And IsNumeric(txtX) Then RaiseEvent ConvPTtoMGRS(txtX, txtY)
    End If

End Sub


Private Sub txtMGRS_LostFocus()

    If Len(txtMGRS) > 3 And Not Len(txtY) > 0 And Not Len(txtX) > 0 Then
        txtX.Text = 0
        txtY.Text = 0
        txtMGRS.Text = Replace(txtMGRS.Text, " ", "")
        RaiseEvent ConvMGRStoPT(txtMGRS, txtX, txtY)
    End If
    
End Sub
