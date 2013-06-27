VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmAddWhat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "What?"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   5970
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5010
      _cx             =   8837
      _cy             =   10530
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
      Begin C1SizerLibCtl.C1Elastic elCommands 
         Height          =   600
         Left            =   180
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   5265
         Width           =   4560
         _cx             =   8043
         _cy             =   1058
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
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            Height          =   465
            Left            =   3105
            TabIndex        =   53
            Top             =   90
            Width           =   915
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
            Height          =   465
            Left            =   1980
            TabIndex        =   52
            Top             =   90
            Width           =   1140
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   465
            Left            =   945
            TabIndex        =   51
            Top             =   90
            Width           =   1050
         End
      End
      Begin C1SizerLibCtl.C1Tab tabWhat 
         Height          =   5100
         Left            =   135
         TabIndex        =   1
         Top             =   135
         Width           =   4830
         _cx             =   8520
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
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "Section/Cluster|Locations|Details|Beneficiaries|Implementing Partners|funding"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   3
         BorderWidth     =   0
         BoldCurrent     =   0   'False
         DogEars         =   -1  'True
         MultiRow        =   -1  'True
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
         Begin C1SizerLibCtl.C1Elastic elFunding 
            Height          =   4440
            Left            =   6675
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   615
            Width           =   4545
            _cx             =   8017
            _cy             =   7832
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
            Begin VB.ComboBox ComReportedFTS 
               Height          =   315
               Left            =   1980
               TabIndex        =   48
               Text            =   "ReportedFTS"
               Top             =   2970
               Width           =   1275
            End
            Begin VB.ComboBox ComFundingStatus 
               Height          =   315
               Left            =   1665
               TabIndex        =   47
               Text            =   "Funding Status"
               Top             =   2160
               Width           =   1275
            End
            Begin VB.ComboBox ComFundingType 
               Height          =   315
               Left            =   1485
               TabIndex        =   46
               Text            =   "FundingType"
               Top             =   1485
               Width           =   1590
            End
            Begin VB.TextBox txtCurrency 
               Height          =   375
               Left            =   1710
               TabIndex        =   45
               Text            =   "Currency"
               Top             =   675
               Width           =   1365
            End
            Begin VB.TextBox txtFundingAmount 
               Height          =   285
               Left            =   1755
               TabIndex        =   44
               Text            =   "Funding Amount"
               Top             =   315
               Width           =   1365
            End
            Begin VB.Label lblReportedTo 
               AutoSize        =   -1  'True
               Caption         =   "Reported To FTS:"
               Height          =   195
               Left            =   450
               TabIndex        =   43
               Top             =   3150
               Width           =   1290
            End
            Begin VB.Label lblFundingStatus 
               AutoSize        =   -1  'True
               Caption         =   "Funding Status:"
               Height          =   195
               Left            =   450
               TabIndex        =   42
               Top             =   2295
               Width           =   1110
            End
            Begin VB.Label lblFundingType 
               AutoSize        =   -1  'True
               Caption         =   "Funding Type:"
               Height          =   195
               Left            =   270
               TabIndex        =   41
               Top             =   1530
               Width           =   1020
            End
            Begin VB.Label lblFundingCurrency 
               AutoSize        =   -1  'True
               Caption         =   "Funding Currency:"
               Height          =   195
               Left            =   270
               TabIndex        =   40
               Top             =   855
               Width           =   1290
            End
            Begin VB.Label lblAmountFunded 
               AutoSize        =   -1  'True
               Caption         =   "Amount Funded:"
               Height          =   195
               Left            =   315
               TabIndex        =   39
               Top             =   315
               Width           =   1170
            End
         End
         Begin C1SizerLibCtl.C1Elastic elImplementingPartner 
            Height          =   4440
            Left            =   6375
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   615
            Width           =   4545
            _cx             =   8017
            _cy             =   7832
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
            Begin VB.CommandButton cmdAddOrganisation 
               Caption         =   "Add Organisation"
               Height          =   600
               Left            =   540
               TabIndex        =   49
               Top             =   3330
               Width           =   1140
            End
            Begin VB.ListBox lstImplementingPartner 
               Height          =   2400
               Left            =   540
               TabIndex        =   38
               Top             =   855
               Width           =   3255
            End
            Begin VB.Label lblImplementingPartners 
               Caption         =   "Implementing Partners:"
               Height          =   510
               Left            =   270
               TabIndex        =   37
               Top             =   315
               Width           =   1860
            End
         End
         Begin C1SizerLibCtl.C1Elastic elBeneficiaries 
            Height          =   4440
            Left            =   6075
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   615
            Width           =   4545
            _cx             =   8017
            _cy             =   7832
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
            Begin VB.TextBox txtSecondaryBeneficiaries 
               Height          =   330
               Left            =   1440
               TabIndex        =   32
               Text            =   "#SecondaryBeneficiaries"
               Top             =   2160
               Width           =   1320
            End
            Begin VB.ListBox lst2ndBeneficiaries 
               Height          =   840
               Left            =   1440
               TabIndex        =   31
               Top             =   1305
               Width           =   2490
            End
            Begin VB.TextBox txtPrimaryBeneficiaries 
               Height          =   285
               Left            =   1485
               TabIndex        =   30
               Text            =   "# Primary Beneficiaries:"
               Top             =   855
               Width           =   1185
            End
            Begin VB.ComboBox ComPrimaryBeneficiary 
               Height          =   315
               Left            =   1440
               TabIndex        =   29
               Text            =   "PrimaryBeneficiary"
               Top             =   450
               Width           =   2625
            End
            Begin VB.Label lblNumSecondaryBeneficiary 
               AutoSize        =   -1  'True
               Caption         =   "# Secondary Beneficiary:"
               Height          =   195
               Left            =   180
               TabIndex        =   36
               Top             =   2250
               Width           =   1785
            End
            Begin VB.Label lblSecondaryBeneficiary 
               AutoSize        =   -1  'True
               Caption         =   "Secondary Beneficiary:"
               Height          =   195
               Left            =   180
               TabIndex        =   35
               Top             =   1350
               Width           =   1635
            End
            Begin VB.Label lblNumPrimaryBeneficiary 
               AutoSize        =   -1  'True
               Caption         =   "# Primary Beneficiary:"
               Height          =   195
               Left            =   135
               TabIndex        =   34
               Top             =   855
               Width           =   1530
            End
            Begin VB.Label lblPrimaryBeneficiary 
               AutoSize        =   -1  'True
               Caption         =   "Primary Beneficiary:"
               Height          =   195
               Left            =   90
               TabIndex        =   33
               Top             =   450
               Width           =   1380
            End
         End
         Begin C1SizerLibCtl.C1Elastic elDetails 
            Height          =   4440
            Left            =   5775
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   615
            Width           =   4545
            _cx             =   8017
            _cy             =   7832
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
            Begin VB.TextBox txtCapProject 
               Height          =   285
               Left            =   720
               TabIndex        =   28
               Text            =   "Cap Project #"
               Top             =   4095
               Width           =   2715
            End
            Begin VB.ComboBox ComProjectTheme 
               Height          =   315
               Left            =   720
               TabIndex        =   27
               Text            =   "ProjectTheme"
               Top             =   3555
               Width           =   2760
            End
            Begin VB.ComboBox ComProjectStatus 
               Height          =   315
               Left            =   720
               TabIndex        =   26
               Text            =   "Project Status"
               Top             =   3015
               Width           =   2760
            End
            Begin VB.ComboBox ComProjectType 
               Height          =   315
               Left            =   720
               TabIndex        =   25
               Text            =   "Project type"
               Top             =   2475
               Width           =   2760
            End
            Begin VB.TextBox txtProjectObjective 
               Height          =   600
               Left            =   720
               TabIndex        =   24
               Text            =   "ProjectObjective"
               Top             =   1620
               Width           =   2940
            End
            Begin VB.TextBox txtProjectDescription 
               Height          =   555
               Left            =   720
               TabIndex        =   23
               Text            =   "ProjectDescription"
               Top             =   810
               Width           =   2940
            End
            Begin VB.TextBox txtProjectTitle 
               Height          =   285
               Left            =   720
               TabIndex        =   22
               Text            =   "Project Title"
               Top             =   270
               Width           =   1995
            End
            Begin VB.Label lblCapProject 
               Caption         =   "Cap Project #:"
               Height          =   195
               Left            =   720
               TabIndex        =   60
               Top             =   3870
               Width           =   1275
            End
            Begin VB.Label lblProjectTheme 
               AutoSize        =   -1  'True
               Caption         =   "Project Theme:"
               Height          =   195
               Left            =   720
               TabIndex        =   59
               Top             =   3330
               Width           =   1080
            End
            Begin VB.Label lblProjectStatus 
               AutoSize        =   -1  'True
               Caption         =   "Project Status:"
               Height          =   195
               Left            =   720
               TabIndex        =   58
               Top             =   2790
               Width           =   1035
            End
            Begin VB.Label lblProjectType 
               AutoSize        =   -1  'True
               Caption         =   "Project Type:"
               Height          =   195
               Left            =   720
               TabIndex        =   57
               Top             =   2250
               Width           =   945
            End
            Begin VB.Label lblProjectObjective 
               AutoSize        =   -1  'True
               Caption         =   "Project Objective:"
               Height          =   195
               Left            =   720
               TabIndex        =   56
               Top             =   1395
               Width           =   1260
            End
            Begin VB.Label lblProjectDescription 
               AutoSize        =   -1  'True
               Caption         =   "Project Description"
               Height          =   195
               Left            =   720
               TabIndex        =   55
               Top             =   585
               Width           =   1335
            End
            Begin VB.Label lblProjectTitle 
               AutoSize        =   -1  'True
               Caption         =   "Project Title:"
               Height          =   195
               Left            =   675
               TabIndex        =   54
               Top             =   45
               Width           =   885
            End
         End
         Begin C1SizerLibCtl.C1Elastic elLocation 
            Height          =   4440
            Left            =   5475
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   615
            Width           =   4545
            _cx             =   8017
            _cy             =   7832
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
            Begin VB.Frame FraLocator 
               Caption         =   "Locator"
               Height          =   2580
               Left            =   495
               TabIndex        =   15
               Top             =   315
               Width           =   2940
               Begin VB.ComboBox ComDistrict 
                  Height          =   315
                  Left            =   270
                  TabIndex        =   18
                  Text            =   "District"
                  Top             =   1980
                  Width           =   2265
               End
               Begin VB.ComboBox ComProvince 
                  Height          =   315
                  Left            =   270
                  TabIndex        =   17
                  Text            =   "Province"
                  Top             =   1170
                  Width           =   2265
               End
               Begin VB.ComboBox ComCountry 
                  Height          =   315
                  Left            =   225
                  TabIndex        =   16
                  Text            =   "Country"
                  Top             =   450
                  Width           =   2310
               End
               Begin VB.Label lblDistrict 
                  AutoSize        =   -1  'True
                  Caption         =   "District:"
                  Height          =   195
                  Left            =   315
                  TabIndex        =   21
                  Top             =   1665
                  Width           =   525
               End
               Begin VB.Label lblProvince 
                  AutoSize        =   -1  'True
                  Caption         =   "Province:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   20
                  Top             =   810
                  Width           =   675
               End
               Begin VB.Label lblCountry 
                  AutoSize        =   -1  'True
                  Caption         =   "Country:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   19
                  Top             =   225
                  Width           =   585
               End
            End
            Begin VB.CommandButton cmdLocate 
               Caption         =   "Locate"
               Height          =   420
               Left            =   270
               TabIndex        =   14
               Top             =   3600
               Width           =   1275
            End
         End
         Begin C1SizerLibCtl.C1Elastic elCluster 
            Height          =   4440
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   615
            Width           =   4545
            _cx             =   8017
            _cy             =   7832
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
            Begin VB.ListBox lstSubSector 
               Height          =   2010
               Left            =   360
               TabIndex        =   12
               Top             =   1755
               Width           =   3435
            End
            Begin VB.ComboBox ComCluster 
               Height          =   315
               Left            =   270
               TabIndex        =   11
               Text            =   "Cluster"
               Top             =   990
               Width           =   2445
            End
            Begin VB.ComboBox ComOrganizationName 
               Height          =   315
               Left            =   270
               TabIndex        =   9
               Text            =   "Organization Name"
               Top             =   405
               Width           =   2490
            End
            Begin VB.Label lblSubSectors 
               AutoSize        =   -1  'True
               Caption         =   "Sub-Sectors:"
               Height          =   195
               Left            =   315
               TabIndex        =   13
               Top             =   1485
               Width           =   915
            End
            Begin VB.Label lblSectorCluster 
               AutoSize        =   -1  'True
               Caption         =   "Sector/Cluster:"
               Height          =   195
               Left            =   270
               TabIndex        =   10
               Top             =   765
               Width           =   1065
            End
            Begin VB.Label lblOrganizationName 
               AutoSize        =   -1  'True
               Caption         =   "Organization Name:"
               Height          =   195
               Left            =   225
               TabIndex        =   8
               Top             =   180
               Width           =   1395
            End
         End
      End
   End
End
Attribute VB_Name = "frmAddWhat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cn As ADODB.Connection

Public Sub Init(cn As Connection)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        Dim rs As New ADODB.Recordset

100     Set m_cn = cn
    
102     rs.Open "SELECT name, id FROM 1organisation", m_cn, adOpenForwardOnly, adLockReadOnly
    
104     ComOrganizationName.Clear
106     ComCluster.Clear
108     ComCountry.Clear
110     ComDistrict.Clear
112     ComFundingStatus.Clear
114     ComFundingType.Clear
116     ComPrimaryBeneficiary.Clear
118     ComProjectStatus.Clear
120     ComProjectTheme.Clear
122     ComProjectType.Clear
124     ComProvince.Clear
126     ComReportedFTS.Clear
    
128    SafeMoveFirst rs
    
130     Do While Not rs.EOF

132         With rs.Fields
134             ComOrganizationName.AddItem .Item("name").Value
                'ComOrgType.itemData(ComOrgType.ListCount - 1) = CLng(.Item("id").value)
            End With

136         rs.MoveNext
        Loop
    
138     Set rs = New ADODB.Recordset
    
140     rs.Open "SELECT name, id FROM 1sector", m_cn, adOpenForwardOnly, adLockReadOnly
    
142     Do While Not rs.EOF

144         With rs.Fields
146             ComCluster.AddItem .Item("name").Value
                'ComOfficestatus.itemData(ComOfficestatus.ListCount - 1) = .Item("id").value
            End With

148         rs.MoveNext
        Loop
    
150     Set rs = New ADODB.Recordset
    
152     rs.Open "SELECT name, id FROM 1country", m_cn, adOpenForwardOnly, adLockReadOnly
    
154     Do While Not rs.EOF

156         With rs.Fields
158             ComCountry.AddItem .Item("name").Value
                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
            End With

160         rs.MoveNext
        Loop
    
162     Set rs = New ADODB.Recordset
    
164     rs.Open "SELECT name, id FROM 1admin2Names", m_cn, adOpenForwardOnly, adLockReadOnly
    
166     Do While Not rs.EOF

168         With rs.Fields
170             ComDistrict.AddItem .Item("name").Value
                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
            End With

172         rs.MoveNext
        Loop
    
174     Set rs = New ADODB.Recordset
    
176     rs.Open "SELECT name, id FROM 1fundingStatus", m_cn, adOpenForwardOnly, adLockReadOnly
    
178     Do While Not rs.EOF

180         With rs.Fields
182             ComFundingStatus.AddItem .Item("name").Value
                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
            End With

184         rs.MoveNext
        Loop
    
186     Set rs = New ADODB.Recordset
    
188     rs.Open "SELECT name, id FROM 1fundingType", m_cn, adOpenForwardOnly, adLockReadOnly
    
190     Do While Not rs.EOF

192         With rs.Fields
194             ComFundingType.AddItem .Item("name").Value
                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
            End With

196         rs.MoveNext
        Loop
    
198     Set rs = New ADODB.Recordset
    
200     rs.Open "SELECT name, id FROM 1beneficiary", m_cn, adOpenForwardOnly, adLockReadOnly
    
202     Do While Not rs.EOF

204         With rs.Fields
206             ComPrimaryBeneficiary.AddItem .Item("name").Value
                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
            End With

208         rs.MoveNext
        Loop
    
210     Set rs = New ADODB.Recordset
    
212     rs.Open "SELECT name, id FROM 1projectStatus", m_cn, adOpenForwardOnly, adLockReadOnly
    
214     Do While Not rs.EOF

216         With rs.Fields
218             ComProjectStatus.AddItem .Item("name").Value
                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
            End With

220         rs.MoveNext
        Loop
    
        '    Set RS = New ADODB.Recordset
        '
        '    RS.Open "SELECT name, id FROM 1projectStatus", m_cn, adOpenForwardOnly, adLockReadOnly
        '
        '    Do While Not RS.EOF
        '        With RS.Fields
        '            ComProjectTheme.AddItem .Item("name").value
        '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
        '        End With
        '
        '        RS.MoveNext
        '    Loop
    
222     Set rs = New ADODB.Recordset
    
224     rs.Open "SELECT name, id FROM 1projectType", m_cn, adOpenForwardOnly, adLockReadOnly
    
226     Do While Not rs.EOF

228         With rs.Fields
230             ComProjectType.AddItem .Item("name").Value
                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
            End With

232         rs.MoveNext
        Loop
    
234     Set rs = New ADODB.Recordset
    
236     rs.Open "SELECT name, id FROM 1admin1Names", m_cn, adOpenForwardOnly, adLockReadOnly
    
238     Do While Not rs.EOF

240         With rs.Fields
242             ComProvince.AddItem .Item("name").Value
                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
            End With

244         rs.MoveNext
        Loop
    
        '    Set RS = New ADODB.Recordset
        '
        '    RS.Open "SELECT name, id FROM 1admin1Names", m_cn, adOpenForwardOnly, adLockReadOnly
        '
        '    Do While Not RS.EOF
        '        With RS.Fields
        '            ComReportedFTS.AddItem .Item("name").value
        '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
        '        End With
        '
        '        RS.MoveNext
        '    Loop
        '
    
246     txtCapProject.Text = ""
248     txtCurrency.Text = ""
250     txtFundingAmount.Text = ""
252     txtPrimaryBeneficiaries.Text = ""
254     txtProjectDescription.Text = ""
256     txtProjectObjective.Text = ""
258     txtProjectTitle.Text = ""
260     txtSecondaryBeneficiaries.Text = ""
        
        On Error Resume Next
262     ComCluster.ListIndex = 0
264     ComCountry.ListIndex = 0
266     ComDistrict.ListIndex = 0
268     ComFundingStatus.ListIndex = 0
270     ComFundingType.ListIndex = 0
272     ComOrganizationName.ListIndex = 0
274     ComPrimaryBeneficiary.ListIndex = 0
276     ComProjectStatus.ListIndex = 0
278     ComProjectTheme.ListIndex = 0
280     ComProjectType.ListIndex = 0
282     ComProvince.ListIndex = 0
284     ComReportedFTS.ListIndex = 0
    
286     tabWhat.TabVisible(0) = True
    
288     g_RSAppSettings.Find "SettingName = 'w3WHATLocation'"
    
290     tabWhat.TabVisible(1) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = "1", True, False)
    
292     g_RSAppSettings.Find "SettingName = 'w3WHATDetails'"
294     tabWhat.TabVisible(2) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = "1", True, False)
    
296     g_RSAppSettings.Find "SettingName = 'w3WHATBeneficiaries'"
298     tabWhat.TabVisible(3) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = "1", True, False)
    
300     g_RSAppSettings.Find "SettingName = 'w3WHATImplementingPartners'"
302     tabWhat.TabVisible(4) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = "1", True, False)
    
304     g_RSAppSettings.Find "SettingName = 'w3WHATFunding'"
306     tabWhat.TabVisible(5) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = "1", True, False)
    
    
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhat.init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAdd_Click()
        '<EhHeader>
        On Error GoTo cmdAdd_Click_Err
        '</EhHeader>
        On Error Resume Next
100     ComCluster.ListIndex = 0
102     ComCountry.ListIndex = 0
104     ComDistrict.ListIndex = 0
106     ComFundingStatus.ListIndex = 0
108     ComFundingType.ListIndex = 0
110     ComOrganizationName.ListIndex = 0
112     ComPrimaryBeneficiary.ListIndex = 0
114     ComProjectStatus.ListIndex = 0
116     ComProjectTheme.ListIndex = 0
118     ComProjectType.ListIndex = 0
120     ComProvince.ListIndex = 0
122     ComReportedFTS.ListIndex = 0
        '<EhFooter>
        Exit Sub

cmdAdd_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhat.cmdAdd_Click " & _
               "at line " & Erl
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

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhat.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
