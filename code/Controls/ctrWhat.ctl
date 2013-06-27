VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.UserControl ctrWhat 
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11760
   ScaleHeight     =   6180
   ScaleWidth      =   11760
   Begin MSAdodcLib.Adodc AdodcWhere 
      Height          =   330
      Left            =   1260
      Top             =   5760
      Visible         =   0   'False
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Where"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin XpressEditorsLibCtl.dxPickEdit dxPickMoveData 
      Height          =   315
      Left            =   180
      OleObjectBlob   =   "ctrWhat.ctx":0000
      TabIndex        =   21
      Top             =   5130
      Width           =   1950
   End
   Begin MSAdodcLib.Adodc AdodcOrgLookUp 
      Height          =   420
      Left            =   2880
      Top             =   5625
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   741
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdodcOrgLookUp"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   495
      Top             =   5670
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "what"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   6180
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11760
      _cx             =   20743
      _cy             =   10901
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
      GridRows        =   1
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"ctrWhat.ctx":0411
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab tabWhat 
         Height          =   6180
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   11760
         _cx             =   20743
         _cy             =   10901
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
         Caption         =   "Project details|Beneficiaries And Implementing partners|Funding and Areas Of Operation"
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
         Begin C1SizerLibCtl.C1Elastic elFunding 
            Height          =   5805
            Left            =   12705
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   11670
            _cx             =   20585
            _cy             =   10239
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
            Begin VB.Frame FraLocator 
               Caption         =   "Areas Of Project Operations"
               Height          =   2490
               Left            =   5670
               TabIndex        =   50
               Top             =   135
               Width           =   5550
               Begin DXDBGRIDLibCtl.dxDBGrid dxDBAdminLocator 
                  Height          =   1995
                  Left            =   135
                  OleObjectBlob   =   "ctrWhat.ctx":0445
                  TabIndex        =   51
                  Top             =   270
                  Width           =   5190
               End
            End
            Begin VB.CommandButton cmdLocate 
               Caption         =   "add Location"
               Height          =   420
               Left            =   5670
               TabIndex        =   49
               Top             =   2880
               Width           =   1275
            End
            Begin MSAdodcLib.Adodc AdodcfundingStatus 
               Height          =   330
               Left            =   540
               Top             =   4140
               Visible         =   0   'False
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   582
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   "fundingStatus"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin MSAdodcLib.Adodc AdodcFundingType 
               Height          =   330
               Left            =   540
               Top             =   3825
               Visible         =   0   'False
               Width           =   2490
               _ExtentX        =   4392
               _ExtentY        =   582
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   "FundingType"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpFundingStatus 
               DataField       =   "fundingStatusId"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   90
               OleObjectBlob   =   "ctrWhat.ctx":3482
               TabIndex        =   27
               Top             =   2520
               Width           =   5190
            End
            Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpFundingType 
               DataField       =   "fundingTypeId"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   90
               OleObjectBlob   =   "ctrWhat.ctx":366B
               TabIndex        =   26
               Top             =   1800
               Width           =   5190
            End
            Begin XpressEditorsLibCtl.dxTextEdit dxedtFundingCurrency 
               DataField       =   "fundingCurrency"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   90
               OleObjectBlob   =   "ctrWhat.ctx":3854
               TabIndex        =   25
               Top             =   1125
               Width           =   5190
            End
            Begin XpressEditorsLibCtl.dxTextEdit dxedtAmountFunded 
               DataField       =   "fundingAmount"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   90
               OleObjectBlob   =   "ctrWhat.ctx":38AC
               TabIndex        =   24
               Top             =   495
               Width           =   5190
            End
            Begin VB.ComboBox ComReportedFTS 
               Height          =   315
               Left            =   90
               TabIndex        =   3
               Text            =   "ReportedFTS"
               Top             =   3240
               Width           =   5235
            End
            Begin XpressEditorsLibCtl.dxTextEdit dxedtProjectId 
               DataField       =   "id"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   5850
               OleObjectBlob   =   "ctrWhat.ctx":3904
               TabIndex        =   48
               Top             =   3645
               Visible         =   0   'False
               Width           =   645
            End
            Begin VB.Label lblAmountFunded 
               AutoSize        =   -1  'True
               Caption         =   "Amount Funded:"
               Height          =   195
               Left            =   90
               TabIndex        =   8
               Top             =   225
               Width           =   1170
            End
            Begin VB.Label lblFundingCurrency 
               AutoSize        =   -1  'True
               Caption         =   "Funding Currency:"
               Height          =   195
               Left            =   90
               TabIndex        =   7
               Top             =   900
               Width           =   1290
            End
            Begin VB.Label lblFundingType 
               AutoSize        =   -1  'True
               Caption         =   "Funding Type:"
               Height          =   195
               Left            =   90
               TabIndex        =   6
               Top             =   1530
               Width           =   1020
            End
            Begin VB.Label lblFundingStatus 
               AutoSize        =   -1  'True
               Caption         =   "Funding Status:"
               Height          =   195
               Left            =   90
               TabIndex        =   5
               Top             =   2295
               Width           =   1110
            End
            Begin VB.Label lblReportedTo 
               AutoSize        =   -1  'True
               Caption         =   "Reported To FTS:"
               Height          =   195
               Left            =   90
               TabIndex        =   4
               Top             =   2970
               Width           =   1290
            End
         End
         Begin C1SizerLibCtl.C1Elastic elBeneficiaries 
            Height          =   5805
            Left            =   12405
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   330
            Width           =   11670
            _cx             =   20585
            _cy             =   10239
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
            Begin VB.ListBox lstImplementingPartner 
               Height          =   2400
               Left            =   3915
               TabIndex        =   45
               Top             =   315
               Width           =   5325
            End
            Begin MSAdodcLib.Adodc AdodcBeneficiary 
               Height          =   330
               Left            =   45
               Top             =   3375
               Visible         =   0   'False
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   582
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   "Beneficiary"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpSecondaryBeneficiary 
               DataField       =   "secondarybeneficiaryId"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   135
               OleObjectBlob   =   "ctrWhat.ctx":395C
               TabIndex        =   23
               Top             =   1800
               Width           =   3120
            End
            Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpPrimaryBeneficiary 
               DataField       =   "primarybeneficiaryId"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   135
               OleObjectBlob   =   "ctrWhat.ctx":3B45
               TabIndex        =   22
               Top             =   360
               Width           =   3210
            End
            Begin VB.TextBox txtPrimaryBeneficiaries 
               DataField       =   "numOfPrimarybeneficiary"
               DataSource      =   "AdodcWhere"
               Height          =   285
               Left            =   135
               TabIndex        =   11
               Text            =   "# Primary Beneficiaries:"
               Top             =   1170
               Width           =   3165
            End
            Begin VB.TextBox txtSecondaryBeneficiaries 
               DataField       =   "numOfSecondarybeneficiary"
               DataSource      =   "AdodcWhere"
               Height          =   330
               Left            =   90
               TabIndex        =   10
               Text            =   "#SecondaryBeneficiaries"
               Top             =   2385
               Width           =   3210
            End
            Begin XpressEditorsLibCtl.dxButtonEdit dxButtonEdit1 
               Height          =   360
               Left            =   3960
               OleObjectBlob   =   "ctrWhat.ctx":3D2E
               TabIndex        =   46
               Top             =   2835
               Width           =   5130
            End
            Begin VB.Label lblImplementingPartners 
               Caption         =   "Implementing Partners:"
               Height          =   510
               Left            =   3915
               TabIndex        =   47
               Top             =   45
               Width           =   1860
            End
            Begin VB.Label lblPrimaryBeneficiary 
               AutoSize        =   -1  'True
               Caption         =   "Primary Beneficiary:"
               Height          =   195
               Left            =   135
               TabIndex        =   15
               Top             =   135
               Width           =   1380
            End
            Begin VB.Label lblNumPrimaryBeneficiary 
               AutoSize        =   -1  'True
               Caption         =   "# Primary Beneficiary:"
               Height          =   195
               Left            =   135
               TabIndex        =   14
               Top             =   855
               Width           =   1530
            End
            Begin VB.Label lblSecondaryBeneficiary 
               AutoSize        =   -1  'True
               Caption         =   "Secondary Beneficiary:"
               Height          =   195
               Left            =   90
               TabIndex        =   13
               Top             =   1575
               Width           =   1635
            End
            Begin VB.Label lblNumSecondaryBeneficiary 
               AutoSize        =   -1  'True
               Caption         =   "# Secondary Beneficiary:"
               Height          =   195
               Left            =   90
               TabIndex        =   12
               Top             =   2160
               Width           =   1785
            End
         End
         Begin C1SizerLibCtl.C1Elastic elCluster 
            Height          =   5805
            Left            =   45
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   330
            Width           =   11670
            _cx             =   20585
            _cy             =   10239
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
            Begin XpressEditorsLibCtl.dxMemoEdit txtProjectObjective 
               DataField       =   "objective"
               DataSource      =   "AdodcWhere"
               Height          =   1995
               Left            =   5715
               OleObjectBlob   =   "ctrWhat.ctx":4297
               TabIndex        =   31
               Top             =   945
               Width           =   2355
            End
            Begin XpressEditorsLibCtl.dxMemoEdit txtProjectDescription 
               DataField       =   "description"
               DataSource      =   "AdodcWhere"
               Height          =   1950
               Left            =   5715
               OleObjectBlob   =   "ctrWhat.ctx":4393
               TabIndex        =   30
               Top             =   3420
               Width           =   5370
            End
            Begin MSAdodcLib.Adodc AdodcProjStatus 
               Height          =   330
               Left            =   9495
               Top             =   4815
               Visible         =   0   'False
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   582
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   "Status"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin MSAdodcLib.Adodc AdodcProjType 
               Height          =   330
               Left            =   8685
               Top             =   4905
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   582
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   "Type"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin MSAdodcLib.Adodc AdodcSector 
               Height          =   330
               Left            =   3105
               Top             =   5445
               Visible         =   0   'False
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   582
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   "Sector"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpSector 
               DataField       =   "sectorId"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   225
               OleObjectBlob   =   "ctrWhat.ctx":448F
               TabIndex        =   29
               Top             =   990
               Width           =   5370
            End
            Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpOrgName 
               DataField       =   "orgId"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   225
               OleObjectBlob   =   "ctrWhat.ctx":4678
               TabIndex        =   28
               Top             =   405
               Width           =   5370
            End
            Begin VB.ListBox lstSubSector 
               Height          =   3570
               Left            =   270
               TabIndex        =   17
               Top             =   1755
               Width           =   5325
            End
            Begin XpressEditorsLibCtl.dxTextEdit dxedtCap 
               DataField       =   "capNumber"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   8415
               OleObjectBlob   =   "ctrWhat.ctx":4861
               TabIndex        =   32
               Top             =   2700
               Width           =   2625
            End
            Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpProjStatus 
               DataField       =   "projectStatusId"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   8415
               OleObjectBlob   =   "ctrWhat.ctx":48B9
               TabIndex        =   33
               Top             =   1125
               Width           =   2670
            End
            Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpProjectTheme 
               Height          =   315
               Left            =   8415
               OleObjectBlob   =   "ctrWhat.ctx":4AA2
               TabIndex        =   34
               Top             =   1890
               Width           =   2670
            End
            Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpProjType 
               DataField       =   "projectTypeId"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   8415
               OleObjectBlob   =   "ctrWhat.ctx":4C8B
               TabIndex        =   35
               Top             =   405
               Width           =   2715
            End
            Begin XpressEditorsLibCtl.dxTextEdit dxTextEdit1 
               DataField       =   "id"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   7830
               OleObjectBlob   =   "ctrWhat.ctx":4E74
               TabIndex        =   36
               Top             =   540
               Visible         =   0   'False
               Width           =   465
            End
            Begin XpressEditorsLibCtl.dxTextEdit txtProjectTitle 
               DataField       =   "projectTitle"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   5715
               OleObjectBlob   =   "ctrWhat.ctx":4ECC
               TabIndex        =   37
               Top             =   360
               Width           =   2400
            End
            Begin VB.Label lblCapProject 
               Caption         =   "Cap Project #:"
               Height          =   195
               Left            =   8415
               TabIndex        =   44
               Top             =   2385
               Width           =   1275
            End
            Begin VB.Label lblProjectTheme 
               AutoSize        =   -1  'True
               Caption         =   "Project Theme:"
               Height          =   195
               Left            =   8415
               TabIndex        =   43
               Top             =   1650
               Width           =   1080
            End
            Begin VB.Label lblProjectStatus 
               AutoSize        =   -1  'True
               Caption         =   "Project Status:"
               Height          =   195
               Left            =   8415
               TabIndex        =   42
               Top             =   915
               Width           =   1035
            End
            Begin VB.Label lblProjectType 
               AutoSize        =   -1  'True
               Caption         =   "Project Type:"
               Height          =   195
               Left            =   8415
               TabIndex        =   41
               Top             =   180
               Width           =   945
            End
            Begin VB.Label lblProjectObjective 
               AutoSize        =   -1  'True
               Caption         =   "Project Objective:"
               Height          =   195
               Left            =   5670
               TabIndex        =   40
               Top             =   720
               Width           =   1260
            End
            Begin VB.Label lblProjectDescription 
               AutoSize        =   -1  'True
               Caption         =   "Project Description"
               Height          =   195
               Left            =   5715
               TabIndex        =   39
               Top             =   3150
               Width           =   1335
            End
            Begin VB.Label lblProjectTitle 
               AutoSize        =   -1  'True
               Caption         =   "Project Title:"
               Height          =   195
               Left            =   5715
               TabIndex        =   38
               Top             =   135
               Width           =   885
            End
            Begin VB.Label lblOrganizationName 
               AutoSize        =   -1  'True
               Caption         =   "Organization Name:"
               Height          =   195
               Left            =   225
               TabIndex        =   20
               Top             =   180
               Width           =   1395
            End
            Begin VB.Label lblSectorCluster 
               AutoSize        =   -1  'True
               Caption         =   "Sector/Cluster:"
               Height          =   195
               Left            =   270
               TabIndex        =   19
               Top             =   765
               Width           =   1065
            End
            Begin VB.Label lblSubSectors 
               AutoSize        =   -1  'True
               Caption         =   "Sub-Sectors:"
               Height          =   195
               Left            =   225
               TabIndex        =   18
               Top             =   1485
               Width           =   915
            End
         End
      End
   End
   Begin XpressEditorsLibCtl.dxStyleController dxStyleController1 
      Left            =   0
      OleObjectBlob   =   "ctrWhat.ctx":4F24
      Top             =   0
   End
   Begin XpressEditorsLibCtl.dxImageLists dxImageLists1 
      Left            =   0
      OleObjectBlob   =   "ctrWhat.ctx":501E
      Top             =   0
   End
End
Attribute VB_Name = "ctrWhat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'Public DetailClose As Boolean
'Public Event MoveFirstOK()
'Public Event MovePreviousOK()
'Public Event MoveNext()
'Public Event MoveLast()
'Public Event MoveAdd()
'Public Event MoveDelete()
'Public Event MoveUpdate()
'Public Event BOFEOF()
'Public Event BOF(b As Boolean)
'Public Event EOF(b As Boolean)
'Private m_parhwnd As Long
'Private m_cn As adodb.Connection
'Dim RRequery As Boolean
'Private WithEvents M_RS As adodb.Recordset
'Public Event MoveComplete(adReason As adodb.EventReasonEnum, _
'                                pError As adodb.Error, _
'                                adStatus As adodb.EventStatusEnum, _
'                                pRecordset As adodb.Recordset)
''Private Type POINTAPI
''        X As Long
''        Y As Long
''End Type
''Private Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
'
'Private Sub UserControl_Resize()
'    dxPickMoveData.Move 100, UserControl.Height - ((dxPickMoveData.Height * 2) + 100)
'End Sub
'
'Public Sub PosPopupWindow(dx As Long, _
'                          dy As Long)
'
''    With Screen
''        Dim XY As POINTAPI
''        XY.x = dx / .TwipsPerPixelX
''        XY.y = dy / .TwipsPerPixelY
''        ClientToScreen m_parhwnd, XY
''        dx = XY.x * .TwipsPerPixelX
''        dy = XY.y * .TwipsPerPixelY
''    End With
'
'End Sub
'Public Property Get EOF() As Boolean
'    EOF = M_RS.EOF
'End Property
'
'Public Property Get BOF() As Boolean
'    BOF = M_RS.BOF
'End Property
'
'Public Sub MoveFirst()
'    'M_RS.MoveFirst
'        dxPickMoveData_ButtonPressed 0, False
'
'    'If Not Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
'    RaiseEvent MoveFirstOK
'
'End Sub
'
'Public Sub MovePrevious()
'        dxPickMoveData_ButtonPressed 1, False
'
'    'If Not Adodc1.Recordset.BOF Then Adodc1.Recordset.MovePrevious
'    'M_RS.MovePrevious
'
'    RaiseEvent MovePreviousOK
'End Sub
'
'Public Sub MoveNext()
'
'    dxPickMoveData_ButtonPressed 2, False
'
''    If Not Adodc1.Recordset.EOF Then
''        Adodc1.Recordset.MoveNext
''        If Adodc1.Recordset.EOF Then Adodc1.Recordset.MovePrevious
''    End If
'
'    RaiseEvent MoveNext
'End Sub
'
'Public Sub MoveLast()
'    'If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
'        dxPickMoveData_ButtonPressed 3, False
'
'    RaiseEvent MoveLast
'End Sub
'
'Public Sub AddRecord()
'    'Adodc1.Recordset.AddNew
'        dxPickMoveData_ButtonPressed 4, False
'
'    RaiseEvent MoveAdd
'End Sub
'
'Public Sub SubmitNewRecord()
'    Adodc1.Recordset.UpdateBatch adAffectCurrent
'End Sub
'
'Public Sub UndoNew()
'    Dim ap As Variant
'
'    With Adodc1.Recordset
'        ap = .AbsolutePosition
'        .CancelUpdate
'        .Requery
'        .AbsolutePosition = ap
'    End With
'
'End Sub
'
'Public Sub DeleteRecord()
''    Adodc1.Recordset.Delete adAffectCurrent
'
''    If Not Adodc1.Recordset.EOF Then
''        Adodc1.Recordset.MoveNext
''    End If
'    dxPickMoveData_ButtonPressed 5, False
'
'    RaiseEvent MoveDelete
'End Sub
'
'Public Sub UpdateRecord()
'    Adodc1.Recordset.UpdateBatch adAffectCurrent
'    RaiseEvent MoveUpdate
'End Sub
'
'Private Sub dxButtonEdit1_ButtonPressed(ByVal ButtonIndex As Integer, _
'                                        ByVal ButtonDown As Boolean)
''    Dim dx As Long, dy As Long
''
''    If ButtonDown Then
''        If DetailClose Then
''            DetailClose = False
''            Exit Sub
''        End If
''
''        dx = dxButtonEdit1.Left: dy = dxButtonEdit1.Top + dxButtonEdit1.Height
''        PosPopupWindow dx, dy
''        If Not frmImplementingPartner.Initialized Then
''            frmImplementingPartner.Init AdodcWhere.ConnectionString
''        End If
''
''        frmImplementingPartner.Load dx, dy, dxStyleController1
''    Else
''
''        If DetailClose Then DetailClose = False
''    End If
'
'
'frmImplementingPartner.Init AdodcWhere.ConnectionString
'frmImplementingPartner.Show vbModal, m_parhwnd
'End Sub
'
'
'
'Private Sub dxPickMoveData_ButtonPressed(ByVal ButtonIndex As Integer, _
'                                         ByVal ButtonDown As Boolean)
'    'On Error Resume Next
'
'    If Not ButtonDown Then
'
'        With AdodcWhere.Recordset
'
'            Select Case ButtonIndex
'
'                Case 2
'                    .MoveNext
'
'                    If .EOF Then .MoveLast
'
'                Case 1
'                    .MovePrevious
'
'                    If .BOF Then .MoveFirst
'
'                Case 0
'                    .MoveFirst
'
'                Case 3
'                    .MoveLast
'
'                Case 4
'                    .AddNew
'                    .UpDate
'
'                Case 5
'                    .Delete
'                    .MoveNext
'
'                    If .EOF Then .MovePrevious
'
'                Case 6
'                    Dim ap
'                    RRequery = True
'                    ap = .AbsolutePosition
'                    .CancelUpdate
'                    .Requery
'                    RRequery = False
'                    .AbsolutePosition = ap
'            End Select
'
'        End With
'
'    End If
'
'End Sub
'
'
'Private Sub Init2()
'    Dim DBName As String
'
'    On Error Resume Next
'    Randomize
'    'DBName = "C:\OASIS\W3Import.mdb" 'g_sAppPath & "\..\Data\exampleED.mdb"
'
'
'    With AdodcOrgLookUp
'        .ConnectionString = m_Cnn.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName
'        .RecordSource = "1organisation"
'        .Refresh
'    End With
'
''    Set dxLookUpOrg.LookUpDataSource = AdodcOrgLookUp
''    dxLookUpOrg.LookUpKeyFieldName = "id"
''    dxLookUpOrg.LookUpDisplayFieldName = "name"
''    dxLookUpOrg.KeepInSync = True
''    dxLookUpOrg.ListFieldName = "name;acronym"
'    'dxLookUpPlaceType.ListColumns =
'
'End Sub
'
'Private Sub LookUpInit()
'
'    With AdodcOrgLookUp
'        .ConnectionString = m_Cnn.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
'        .RecordSource = "1organisation"
'        .Refresh
'    End With
'
'    With dxLookUpOrgName
'        Set .LookUpDataSource = AdodcOrgLookUp
'        .LookUpKeyFieldName = "id"
'        .LookUpDisplayFieldName = "name"
'        .KeepInSync = True
'        .ListFieldName = "name;acronym"
'    End With
'
'    With AdodcProjType
'        .ConnectionString = m_Cnn.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
'        .RecordSource = "1projectType"
'        .Refresh
'    End With
'
'    Set dxLookUpProjType.LookUpDataSource = AdodcProjType
'    dxLookUpProjType.LookUpKeyFieldName = "id"
'    dxLookUpProjType.LookUpDisplayFieldName = "name"
'    dxLookUpProjType.KeepInSync = True
'    dxLookUpProjType.ListFieldName = "name"
'
'    With AdodcProjStatus
'        .ConnectionString = m_Cnn.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
'        .RecordSource = "1projectStatus"
'        .Refresh
'    End With
'
'    Set dxLookUpProjStatus.LookUpDataSource = AdodcProjStatus
'    dxLookUpProjStatus.LookUpKeyFieldName = "id"
'    dxLookUpProjStatus.LookUpDisplayFieldName = "name"
'    dxLookUpProjStatus.KeepInSync = True
'    dxLookUpProjStatus.ListFieldName = "name"
'
'    With AdodcSector
'        .ConnectionString = m_Cnn.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
'        .RecordSource = "1beneficiary"
'        .Refresh
'    End With
'
'    With dxLookUpSector
'
'        Set .LookUpDataSource = AdodcSector
'        .LookUpKeyFieldName = "id"
'        .LookUpDisplayFieldName = "name"
'        .KeepInSync = True
'        .ListFieldName = "name"
'
'    End With
'
'    With AdodcBeneficiary
'        .ConnectionString = m_Cnn.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
'        .RecordSource = "1beneficiary"
'        .Refresh
'    End With
'
'    With dxLookUpPrimaryBeneficiary
'
'        Set .LookUpDataSource = AdodcBeneficiary
'        .LookUpKeyFieldName = "id"
'        .LookUpDisplayFieldName = "name"
'        .KeepInSync = True
'        .ListFieldName = "name"
'
'    End With
'
'    With dxLookUpSecondaryBeneficiary
'
'        Set .LookUpDataSource = AdodcBeneficiary
'        .LookUpKeyFieldName = "id"
'        .LookUpDisplayFieldName = "name"
'        .KeepInSync = True
'        .ListFieldName = "name"
'
'    End With
'
'    With AdodcFundingType
'        .ConnectionString = m_Cnn.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
'        .RecordSource = "1fundingType"
'        .Refresh
'    End With
'
'    With dxLookUpFundingType
'
'        Set .LookUpDataSource = AdodcFundingType
'        .LookUpKeyFieldName = "id"
'        .LookUpDisplayFieldName = "name"
'        .KeepInSync = True
'        .ListFieldName = "name"
'
'    End With
'
'    With AdodcfundingStatus
'        .ConnectionString = m_Cnn.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
'        .RecordSource = "1fundingStatus"
'        .Refresh
'    End With
'
'    With dxLookUpFundingStatus
'
'        Set .LookUpDataSource = AdodcfundingStatus
'        .LookUpKeyFieldName = "id"
'        .LookUpDisplayFieldName = "name"
'        .KeepInSync = True
'        .ListFieldName = "name"
'
'    End With
'
'    With AdodcWhere
'        .ConnectionString = m_Cnn.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
'        .RecordSource = "projects"
'        .Refresh
'    End With
'
'End Sub
'
'Public Sub Init(CN As Connection, _
'                parhwnd As Long, Optional bShowInternalNavigator As Boolean = True)
'        'g_sAppPath & "\..\Data\exampleED.mdb"
'        '<EhHeader>
'        On Error GoTo Init_Err
'        '</EhHeader>
'100     m_parhwnd = parhwnd
'102     LookUpInit
'
'104     With Adodc1
'            '.ConnectionString = cn.ConnectionString
'            '.CommandType = adCmdText
'106         .ConnectionString = m_Cnn.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
'            '.CursorLocation = g_sGlobalCursorLocation
'108         .RecordSource = "project" '"SELECT * FROM qryOrganisation"
'110         .Refresh
'            '.Recordset.MoveFirst
'        End With
'
'        dxPickMoveData.Visible = bShowInternalNavigator
'        'Init2
'
'        Dim RS As New adodb.Recordset
'
'112     Set m_cn = CN
'
'        '    RS.Open "SELECT name, id FROM 1organisation", m_cn, adOpenForwardOnly, adLockReadOnly
'        '    Set M_RS = New ADODB.Recordset
'        '    M_RS.Open "SELECT name, id FROM 1organisation", m_cn, adOpenDynamic, adLockOptimistic
'        '
'        '    Set Adodc1.Recordset = M_RS
'        '
'        '    ComOrganizationName.Clear
'        '    ComCluster.Clear
'        '    ComCountry.Clear
'        '    ComDistrict.Clear
'        '    ComFundingStatus.Clear
'        '    ComFundingType.Clear
'        '    ComPrimaryBeneficiary.Clear
'        '    ComProjectStatus.Clear
'        '    ComProjectTheme.Clear
'        '    ComProjectType.Clear
'        '    ComProvince.Clear
'        '    ComReportedFTS.Clear
'        '
'        '    RS.MoveFirst
'        '
'        '    Do While Not RS.EOF
'        '
'        '        With RS.Fields
'        '            ComOrganizationName.AddItem .Item("name").value
'        '            peOrgName.Items.Add .Item("name").value
'        '            'ComOrgType.itemData(ComOrgType.ListCount - 1) = CLng(.Item("id").value)
'        '        End With
'        '
'        '        RS.MoveNext
'        '    Loop
'        '
'        '    Set RS = New ADODB.Recordset
'        '
'        '    RS.Open "SELECT name, id FROM 1sector", m_cn, adOpenForwardOnly, adLockReadOnly
'        '
'        '    Do While Not RS.EOF
'        '
'        '        With RS.Fields
'        '            ComCluster.AddItem .Item("name").value
'        '            'ComOfficestatus.itemData(ComOfficestatus.ListCount - 1) = .Item("id").value
'        '        End With
'        '
'        '        RS.MoveNext
'        '    Loop
'        '
'        '    Set RS = New ADODB.Recordset
'        '
'        '    RS.Open "SELECT name, id FROM 1country", m_cn, adOpenForwardOnly, adLockReadOnly
'        '
'        '    Do While Not RS.EOF
'        '
'        '        With RS.Fields
'        '            ComCountry.AddItem .Item("name").value
'        '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
'        '        End With
'        '
'        '        RS.MoveNext
'        '    Loop
'        '
'        '    Set RS = New ADODB.Recordset
'        '
'        '    RS.Open "SELECT name, id FROM 1admin2Names", m_cn, adOpenForwardOnly, adLockReadOnly
'        '
'        '    Do While Not RS.EOF
'        '
'        '        With RS.Fields
'        '            ComDistrict.AddItem .Item("name").value
'        '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
'        '        End With
'        '
'        '        RS.MoveNext
'        '    Loop
'        '
'        '    Set RS = New ADODB.Recordset
'        '
'        '    RS.Open "SELECT name, id FROM 1fundingStatus", m_cn, adOpenForwardOnly, adLockReadOnly
'        '
'        '    Do While Not RS.EOF
'        '
'        '        With RS.Fields
'        '            ComFundingStatus.AddItem .Item("name").value
'        '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
'        '        End With
'        '
'        '        RS.MoveNext
'        '    Loop
'        '
'        '    Set RS = New ADODB.Recordset
'        '
'        '    RS.Open "SELECT name, id FROM 1fundingType", m_cn, adOpenForwardOnly, adLockReadOnly
'        '
'        '    Do While Not RS.EOF
'        '
'        '        With RS.Fields
'        '            ComFundingType.AddItem .Item("name").value
'        '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
'        '        End With
'        '
'        '        RS.MoveNext
'        '    Loop
'        '
'        '    Set RS = New ADODB.Recordset
'        '
'        '    RS.Open "SELECT name, id FROM 1beneficiary", m_cn, adOpenForwardOnly, adLockReadOnly
'        '
'        '    Do While Not RS.EOF
'        '
'        '        With RS.Fields
'        '            ComPrimaryBeneficiary.AddItem .Item("name").value
'        '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
'        '        End With
'        '
'        '        RS.MoveNext
'        '    Loop
'        '
'        '    Set RS = New ADODB.Recordset
'        '
'        '    RS.Open "SELECT name, id FROM 1projectStatus", m_cn, adOpenForwardOnly, adLockReadOnly
'        '
'        '    Do While Not RS.EOF
'        '
'        '        With RS.Fields
'        '            ComProjectStatus.AddItem .Item("name").value
'        '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
'        '        End With
'        '
'        '        RS.MoveNext
'        '    Loop
'        '
'        '    '    Set RS = New ADODB.Recordset
'        '    '
'        '    '    RS.Open "SELECT name, id FROM 1projectStatus", m_cn, adOpenForwardOnly, adLockReadOnly
'        '    '
'        '    '    Do While Not RS.EOF
'        '    '        With RS.Fields
'        '    '            ComProjectTheme.AddItem .Item("name").value
'        '    '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
'        '    '        End With
'        '    '
'        '    '        RS.MoveNext
'        '    '    Loop
'        '
'        '    Set RS = New ADODB.Recordset
'        '
'        '    RS.Open "SELECT name, id FROM 1projectType", m_cn, adOpenForwardOnly, adLockReadOnly
'        '
'        '    Do While Not RS.EOF
'        '
'        '        With RS.Fields
'        '            ComProjectType.AddItem .Item("name").value
'        '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
'        '        End With
'        '
'        '        RS.MoveNext
'        '    Loop
'        '
'        '    Set RS = New ADODB.Recordset
'        '
'        '    RS.Open "SELECT name, id FROM 1admin1Names", m_cn, adOpenForwardOnly, adLockReadOnly
'        '
'        '    Do While Not RS.EOF
'        '
'        '        With RS.Fields
'        '            ComProvince.AddItem .Item("name").value
'        '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
'        '        End With
'        '
'        '        RS.MoveNext
'        '    Loop
'        '
'        '    '    Set RS = New ADODB.Recordset
'        '    '
'        '    '    RS.Open "SELECT name, id FROM 1admin1Names", m_cn, adOpenForwardOnly, adLockReadOnly
'        '    '
'        '    '    Do While Not RS.EOF
'        '    '        With RS.Fields
'        '    '            ComReportedFTS.AddItem .Item("name").value
'        '    '            'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
'        '    '        End With
'        '    '
'        '    '        RS.MoveNext
'        '    '    Loop
'        '    '
'        '
'        '    txtCapProject.Text = ""
'        '    txtCurrency.Text = ""
'        '    txtFundingAmount.Text = ""
'        '    txtPrimaryBeneficiaries.Text = ""
'        '    'txtProjectDescription.DisplayText = ""
'        '    'txtProjectObjective.DisplayText = ""
'        '    'txtProjectTitle.DisplayText = ""
'        '    txtSecondaryBeneficiaries.Text = ""
'        '
'        '    On Error Resume Next
'        '    ComCluster.ListIndex = 0
'        '    ComCountry.ListIndex = 0
'        '    ComDistrict.ListIndex = 0
'        '    ComFundingStatus.ListIndex = 0
'        '    ComFundingType.ListIndex = 0
'        '    ComOrganizationName.ListIndex = 0
'        '    ComPrimaryBeneficiary.ListIndex = 0
'        '    ComProjectStatus.ListIndex = 0
'        '    ComProjectTheme.ListIndex = 0
'        '    ComProjectType.ListIndex = 0
'        '    ComProvince.ListIndex = 0
'        '    ComReportedFTS.ListIndex = 0
'        '
'114     tabWhat.TabVisible(0) = True
'
'116     SafeMoveFirst g_RSAppSettings
'118     g_RSAppSettings.Find "SettingName = 'w3WHATLocation'"
'
'120     tabWhat.TabVisible(1) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = "1", True, False)
'
'122     SafeMoveFirst g_RSAppSettings
'124     g_RSAppSettings.Find "SettingName = 'w3WHATDetails'"
'126     tabWhat.TabVisible(2) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = "1", True, False)
'
'128     With dxDBAdminLocator.Dataset.ADODataset
'130         .ConnectionString = m_Cnn.ConnectionString
'132         .CommandType = cmdTable
'134         .CommandText = "AOP"
'        End With
'
'135     dxDBAdminLocator.Columns(1).LookupColumn.LookupDataset.ADODataset.ConnectionString = m_Cnn
'647     dxDBAdminLocator.Columns(2).LookupColumn.LookupDataset.ADODataset.ConnectionString = m_Cnn
'733     dxDBAdminLocator.Columns(4).LookupColumn.LookupDataset.ADODataset.ConnectionString = m_Cnn
'
'        'dxDBAdminLocator.Dataset.Open
'136     dxDBAdminLocator.Dataset.Active = True
'        'dxDBAdminLocator.Dataset.ADODataset.Requery
'
'        'g_RSAppSettings.MoveFirst
'        'g_RSAppSettings.Find "SettingName = 'w3WHATBeneficiaries'"
'        'tabWhat.TabVisible(3) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = "1", True, False)
'
'        'g_RSAppSettings.MoveFirst
'        'g_RSAppSettings.Find "SettingName = 'w3WHATImplementingPartners'"
'        'tabWhat.TabVisible(4) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = "1", True, False)
'
'        'g_RSAppSettings.MoveFirst
'        'g_RSAppSettings.Find "SettingName = 'w3WHATFunding'"
'        'tabWhat.TabVisible(5) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = "1", True, False)
'
'        '<EhFooter>
'        Exit Sub
'
'Init_Err:
'        MsgBox "Error Num:" & Err.number & " OASISClient.ctrWhat.Init " & Err.Description & " Occured on line:" & Erl
'
'        '</EhFooter>
'End Sub
'
'Private Sub cmdAddWhat_Click()
'    On Error Resume Next
''    ComCluster.ListIndex = 0
''    ComCountry.ListIndex = 0
''    ComDistrict.ListIndex = 0
''    ComFundingStatus.ListIndex = 0
''    ComFundingType.ListIndex = 0
''    ComOrganizationName.ListIndex = 0
''    ComPrimaryBeneficiary.ListIndex = 0
''    ComProjectStatus.ListIndex = 0
''    ComProjectTheme.ListIndex = 0
''    ComProjectType.ListIndex = 0
''    ComProvince.ListIndex = 0
''    ComReportedFTS.ListIndex = 0
'End Sub
'
'Private Sub AdodcWhere_FieldChangeComplete(ByVal cFields As Long, Fields As Variant, ByVal pError As adodb.Error, adStatus As adodb.EventStatusEnum, ByVal pRecordset As adodb.Recordset)
'  dxPickMoveData.Buttons(6).Enabled = True
'End Sub
'
'Private Sub AdodcWhere_MoveComplete(ByVal adReason As adodb.EventReasonEnum, _
'                                    ByVal pError As adodb.Error, _
'                                    adStatus As adodb.EventStatusEnum, _
'                                    ByVal pRecordset As adodb.Recordset)
'    On Error Resume Next
'
'    Exit Sub
'
'    If RRequery Then Exit Sub
'
'    'dxDBAdminLocator.Dataset.ADODataset.CommandType = cmdText
'    'dxDBAdminLocator.Dataset.ADODataset.CommandText = "SELECT * FROM AOP WHERE projectId = " & dxedtProjectId.EditValue
'    'dxDBAdminLocator.Dataset.ADODataset.Requery
'    'dxDBAdminLocator.Dataset.Refresh
'
'    With AdodcWhere.Recordset
'
'        dxPickMoveData.Buttons(6).Enabled = False
'        dxPickMoveData.Buttons(5).Enabled = True
'        dxPickMoveData.Buttons(2).Enabled = .AbsolutePosition <> .RecordCount
'        dxPickMoveData.Buttons(1).Enabled = .AbsolutePosition <> 1
'
'        If .AbsolutePosition = 0 Then dxPickMoveData.Buttons(2).Enabled = False
'        If .RecordCount < 1 Then dxPickMoveData.Buttons(1).Enabled = False
'        If .EOF And .BOF Then
'            dxPickMoveData.Buttons(2).Enabled = False
'            dxPickMoveData.Buttons(1).Enabled = False
'            dxPickMoveData.Buttons(5).Enabled = False
'        End If
'
'        dxPickMoveData.Buttons(3).Enabled = dxPickMoveData.Buttons(2).Enabled
'        dxPickMoveData.Buttons(0).Enabled = dxPickMoveData.Buttons(1).Enabled
'    End With
'
'End Sub
'
'Private Sub M_RS_MoveComplete(ByVal adReason As adodb.EventReasonEnum, ByVal pError As adodb.Error, adStatus As adodb.EventStatusEnum, ByVal pRecordset As adodb.Recordset)
'    Select Case adReason
'
'        Case EventReasonEnum.adRsnMove
'
'        Case EventReasonEnum.adRsnDelete
'
'        Case EventReasonEnum.adRsnMoveFirst
'
'        Case EventReasonEnum.adRsnMovePrevious
'
'        Case EventReasonEnum.adRsnMoveLast
'
'        Case EventReasonEnum.adRsnMoveNext
'
'        Case EventReasonEnum.adRsnAddNew
'
'        Case EventReasonEnum.adRsnMove
'
'        Case EventReasonEnum.adRsnUpdate
'
'        Case EventReasonEnum.adRsnRequery
'    End Select
'End Sub
'
