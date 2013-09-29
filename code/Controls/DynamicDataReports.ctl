VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Begin VB.UserControl DynamicDataReports 
   ClientHeight    =   8115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   11280
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8115
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11280
      _cx             =   19897
      _cy             =   14314
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
      BackColor       =   12632256
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
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
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   4
      GridCols        =   3
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"DynamicDataReports.ctx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab2 
         Height          =   2520
         Left            =   90
         TabIndex        =   1
         Top             =   5505
         Width           =   11100
         _cx             =   19579
         _cy             =   4445
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
         BackColor       =   128
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "Data View|Export Options|Report Query|Admin"
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
         Flags(3)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   2145
            Left            =   12345
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   330
            Width           =   11010
            _cx             =   19420
            _cy             =   3784
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
            Begin VB.CommandButton cmdRepaint 
               Caption         =   "repaint"
               Height          =   315
               Left            =   8820
               TabIndex        =   36
               Top             =   1500
               Width           =   1335
            End
            Begin VB.CommandButton cmdHalfChatrSize 
               Caption         =   "HalfChatrSize"
               Height          =   495
               Left            =   8700
               TabIndex        =   34
               Top             =   720
               Width           =   1695
            End
            Begin VB.CommandButton cmdREsetChart 
               Caption         =   "REsetChart"
               Height          =   555
               Left            =   5700
               TabIndex        =   33
               Top             =   1140
               Width           =   2775
            End
            Begin VB.CommandButton cmdDoubleChartSize 
               Caption         =   "DoubleChartSize"
               Height          =   615
               Left            =   5700
               TabIndex        =   32
               Top             =   300
               Width           =   2655
            End
            Begin VB.CommandButton cmdMakeChartIn 
               Caption         =   "Make Chart Invisible"
               Height          =   495
               Left            =   2880
               TabIndex        =   31
               Top             =   1080
               Width           =   2055
            End
            Begin VB.CommandButton cmdMakeChart 
               Caption         =   "Make Chart Visible"
               Height          =   435
               Left            =   3000
               TabIndex        =   30
               Top             =   360
               Width           =   1935
            End
            Begin VB.CommandButton cmdSetChart 
               Caption         =   "Set Chart Container"
               Height          =   375
               Left            =   660
               TabIndex        =   29
               Top             =   1080
               Width           =   1875
            End
            Begin VB.CommandButton cmdRefreshChart 
               Caption         =   "Refresh Chart"
               Height          =   435
               Left            =   600
               TabIndex        =   28
               Top             =   300
               Width           =   1815
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1ElasticGrid 
            Height          =   2145
            Left            =   45
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   330
            Width           =   11010
            _cx             =   19420
            _cy             =   3784
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
            BackColor       =   128
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
            GridRows        =   4
            GridCols        =   5
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   0
            FrameColor      =   128
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"DynamicDataReports.ctx":0076
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
               Height          =   1665
               Left            =   30
               OleObjectBlob   =   "DynamicDataReports.ctx":00FF
               TabIndex        =   16
               Top             =   30
               Width           =   10950
            End
            Begin C1SizerLibCtl.C1Elastic C1EDateFrom 
               Height          =   390
               Left            =   5595
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   1725
               Width           =   1185
               _cx             =   2090
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
               BackColor       =   128
               ForeColor       =   16777215
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Date from:"
               Align           =   0
               AutoSizeChildren=   0
               BorderWidth     =   6
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   7
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
            Begin XpressEditorsLibCtl.dxDateEdit dxDateEditFrom 
               Height          =   315
               Left            =   6810
               OleObjectBlob   =   "DynamicDataReports.ctx":0DA7
               TabIndex        =   18
               Top             =   1725
               Width           =   1470
            End
            Begin C1SizerLibCtl.C1Elastic C1EDateTill 
               Height          =   390
               Left            =   8310
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   1725
               Width           =   1185
               _cx             =   2090
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
               BackColor       =   128
               ForeColor       =   16777215
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Date till:"
               Align           =   0
               AutoSizeChildren=   0
               BorderWidth     =   6
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   7
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
            Begin XpressEditorsLibCtl.dxDateEdit dxDateEditTill 
               Height          =   315
               Left            =   9525
               OleObjectBlob   =   "DynamicDataReports.ctx":0E63
               TabIndex        =   20
               Top             =   1725
               Width           =   1455
            End
            Begin C1SizerLibCtl.C1Elastic C1ERecordCount 
               Height          =   390
               Left            =   30
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   1725
               Width           =   5535
               _cx             =   9763
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
               BackColor       =   128
               ForeColor       =   16777215
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Record Count: 0"
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
            Height          =   2145
            Left            =   11745
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   330
            Width           =   11010
            _cx             =   19420
            _cy             =   3784
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
            _GridInfo       =   $"DynamicDataReports.ctx":0F1F
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.CommandButton cmdExport 
               Caption         =   "Export"
               Height          =   345
               Left            =   90
               TabIndex        =   5
               Top             =   1710
               Width           =   10830
            End
            Begin VB.ListBox lstExports 
               BackColor       =   &H80000006&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   1500
               ItemData        =   "DynamicDataReports.ctx":0F62
               Left            =   90
               List            =   "DynamicDataReports.ctx":0F7B
               TabIndex        =   4
               Top             =   90
               Width           =   10830
            End
            Begin VB.PictureBox Picture1 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               Height          =   1560
               Left            =   90
               ScaleHeight     =   1500
               ScaleWidth      =   10770
               TabIndex        =   6
               Top             =   90
               Visible         =   0   'False
               Width           =   10830
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1ElasticQuery 
            Height          =   2145
            Left            =   12045
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   330
            Width           =   11010
            _cx             =   19420
            _cy             =   3784
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
            GridRows        =   8
            GridCols        =   8
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   0
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"DynamicDataReports.ctx":100A
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin XpressEditorsLibCtl.dxMemoEdit dxFilterSQL 
               Height          =   810
               Left            =   5535
               OleObjectBlob   =   "DynamicDataReports.ctx":10E2
               TabIndex        =   23
               Top             =   270
               Width           =   5475
            End
            Begin XpressEditorsLibCtl.dxMemoEdit dxSQLCommand 
               Height          =   1605
               Left            =   0
               OleObjectBlob   =   "DynamicDataReports.ctx":11DE
               TabIndex        =   9
               Top             =   270
               Width           =   5475
            End
            Begin VB.CheckBox chkAutoLoad 
               Caption         =   "Auto Load Report"
               Height          =   210
               Left            =   8295
               TabIndex        =   26
               Top             =   1665
               Width           =   2715
            End
            Begin VB.TextBox txtTxtGroup 
               Height          =   480
               Left            =   5535
               TabIndex        =   25
               Text            =   "txtGroup"
               Top             =   1395
               Width           =   2700
            End
            Begin VB.CommandButton cmdUpdateReport 
               Caption         =   "Update Report"
               Height          =   210
               Left            =   0
               TabIndex        =   13
               Top             =   1935
               Width           =   5475
            End
            Begin VB.CommandButton cmdSaveSQL 
               Caption         =   "Save"
               Height          =   210
               Left            =   5535
               TabIndex        =   12
               Top             =   1935
               Visible         =   0   'False
               Width           =   2700
            End
            Begin VB.CommandButton cmdRemoveChart 
               Caption         =   "Remove"
               Height          =   210
               Left            =   8295
               TabIndex        =   11
               Top             =   1935
               Visible         =   0   'False
               Width           =   2715
            End
            Begin VB.CheckBox chkUseChart 
               Caption         =   "Use Chart"
               Height          =   210
               Left            =   8295
               TabIndex        =   10
               Top             =   1395
               Width           =   2715
            End
            Begin VB.Label lblGroup 
               Caption         =   "Group"
               Height          =   195
               Left            =   5535
               TabIndex        =   24
               Top             =   1140
               Width           =   2700
            End
            Begin VB.Label lblFilterSQL 
               Caption         =   "Filter SQL"
               Height          =   210
               Left            =   5535
               TabIndex        =   22
               Top             =   0
               Width           =   5475
            End
            Begin VB.Label lblSQLCommand 
               Caption         =   "SQL Command"
               Height          =   210
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   5475
            End
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   5355
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   11100
         _cx             =   19579
         _cy             =   9446
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   128
         ForeColor       =   4210816
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   4210816
         Caption         =   "Chart View"
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
         Begin C1SizerLibCtl.C1Elastic C1ElasticChart 
            Height          =   4980
            Left            =   45
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   330
            Width           =   11010
            _cx             =   19420
            _cy             =   8784
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   15.75
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
            Picture         =   "DynamicDataReports.ctx":12DA
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   0
            ChildSpacing    =   3
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   4
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   3
            CaptionStyle    =   4
            ResizeFonts     =   0   'False
            GridRows        =   1
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   0
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"DynamicDataReports.ctx":5041
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin OASISClient.OASISChartingVer2 OASISChartingVer2 
               Height          =   4980
               Left            =   0
               TabIndex        =   35
               Top             =   0
               Width           =   11010
               _ExtentX        =   19420
               _ExtentY        =   8784
            End
         End
      End
   End
   Begin VB.Menu dude 
      Caption         =   "dude"
   End
End
Attribute VB_Name = "DynamicDataReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Private Type DYNAMIC_DATA_DEF_QUERIES
    QueryName As String
    ListIndex As Long
    Command As String
    Group As String
    UseChart As Boolean
    AutoLoadChart As Boolean
    FilterSQL As String
    ChartSettings As String
End Type

Private Type DYNAMIC_DATA_DEF
    Name As String
    Prefix As String
    desc As String
    Queries() As DYNAMIC_DATA_DEF_QUERIES
    ListIndex As Long
    ConnectionString As String
End Type

Private mConn As ADODB.Connection
Private mRS As ADODB.Recordset

Private DDDefs() As DYNAMIC_DATA_DEF
Private DDDefCurrent As DYNAMIC_DATA_DEF
Private DDQueryCurrent As DYNAMIC_DATA_DEF_QUERIES
Private lCurrentQueryIndex As Long
Private bToolbar1 As Boolean
Private bToolbar2 As Boolean
Private listQueries As ListBox
Private cmbDatabases As ListBox
Private cmbFilter As ComboBox
Private listGroups As ListBox

Public Event ExportDone(listQueries As ListBox, Combo1 As ComboBox)

Public Sub GroupClicked()

    ShowQueries '6
    OASISChartingVer2.Visible = False
    C1Tab2.Visible = False

End Sub

Private Sub CreateXLSFromRS(oRS As ADODB.Recordset, _
                            sPath As String)
        '<EhHeader>
        On Error GoTo CreateXLSFromRS_Err
        '</EhHeader>
    
        Dim mstream As New ADODB.Stream
        Dim sText As String
        Dim iCol As Long
        Dim iRow As Long

100     If SafeMoveFirst(oRS) Then
        
102         iCol = 0
104         sText = "<table border=1>"
106         sText = sText & "<tr>"
        
108         Do Until (iCol + 1) > oRS.Fields.Count
110             sText = sText & "<td><b>" & oRS.Fields(iCol).Name & "</b></td>"
112             iCol = iCol + 1
            Loop
        
114         sText = sText & "</tr>"
        
116         Do Until oRS.EOF
        
118             iCol = 0
120             sText = sText & "<tr>"
        
122             Do Until (iCol + 1) > oRS.Fields.Count
124                 sText = sText & "<td>" & oRS.Fields(iCol).value & "</td>"
126                 iCol = iCol + 1
                Loop

128             sText = sText & "</tr>"
            
130             oRS.MoveNext
            Loop
        
132         sText = sText & "</table>"

134         mstream.Type = adTypeText
136         mstream.Open
138         mstream.WriteText sText
140         mstream.SaveToFile sPath
142         mstream.Close

        End If

144     Set mstream = Nothing

        '<EhFooter>
        Exit Sub

CreateXLSFromRS_Err:
        MsgBox "DynamicReports.CreateXLSFromRS_Err (line " & Erl & "): " & Err.Description
        '</EhFooter>
End Sub

Public Sub Init(ByRef lstExports As ListBox, _
                ByRef Combo1 As ListBox, _
                ByRef ComFilter As ComboBox, _
                ByRef listGroup As ListBox)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        
        Dim i As Long
        Dim l As Long
        Dim iDDDefIndex As Long
        Dim oRS As New ADODB.Recordset
        Dim oConn As ADODB.Connection
        Dim bNotAvailable As Boolean
        Dim sConnectionStrings As String
        Dim bIncidentReports As Boolean
        
100     C1Tab1.CurrTab = 0
        C1Tab2.CurrTab = 0
102     OASISChartingVer2.Visible = False

        listGroup.Enabled = False
104     Set listQueries = lstExports
106     Set cmbDatabases = Combo1
        Set cmbFilter = ComFilter
        Set listGroups = listGroup
        
        listQueries.Clear
        cmbDatabases.Clear
        listGroups.Clear
        listGroups.AddItem " -- ALL --"
        listGroups.ListIndex = 0
        
        
        oRS.Open "SELECT * FROM AppSettings WHERE [SettingName] = 'ShowIncidentReports'", m_Cnn, adOpenDynamic, adLockReadOnly
        If Not oRS.EOF Then bIncidentReports = IIf(oRS.Fields("SettingValue1").value = "1", True, False)
        oRS.Close
        
        Set oRS = New ADODB.Recordset
        If bSQLServerInUse Then
            oRS.Open "SELECT * FROM DynamicDataDefs WHERE [EnableReporting] = 'TRUE' ORDER BY cast([Description] as nvarchar(255))", m_Cnn, adOpenDynamic, adLockReadOnly
        Else
108         oRS.Open "SELECT * FROM DynamicDataDefs WHERE [EnableReporting] = TRUE ORDER BY [Description]", m_Cnn, adOpenDynamic, adLockReadOnly
        End If

110     ReDim DDDefs(0)
    
112     l = 0
114     iDDDefIndex = 1
        
116     ReDim Preserve DDDefs(UBound(DDDefs) + 1)
118     DDDefs(UBound(DDDefs)).Name = "Incidents"
120     DDDefs(UBound(DDDefs)).desc = "OASIS Incidents"
122     DDDefs(UBound(DDDefs)).Prefix = "Incidents_"
124     DDDefs(UBound(DDDefs)).ListIndex = iDDDefIndex - 1
126     DDDefs(UBound(DDDefs)).ConnectionString = m_Cnn.ConnectionString
        
128     GetQueries m_Cnn, DDDefs(UBound(DDDefs)), "Incidents_"
             
130     If bIncidentReports And Not UBound(DDDefs(UBound(DDDefs)).Queries) = 0 Then
132         cmbDatabases.AddItem "OASIS Incidents", iDDDefIndex - 1
        Else
134         iDDDefIndex = iDDDefIndex - 1
            ReDim DDDefs(0)
        End If
            
136     Do Until oRS.EOF

138         bNotAvailable = False
140         Set oConn = New ADODB.Connection
142         oConn.ConnectionString = Replace(oRS.Fields("ConnectionString").value, "\data\db\dynamicdata", AppPath & "data\db\dynamicdata", , , vbTextCompare)

            If bSQLServerInUse Then oConn.ConnectionString = g_sGlobalConnectionString
144         oConn.CursorLocation = g_sGlobalCursorLocation
            On Error GoTo NotAvailable
146         oConn.Open
            On Error GoTo Init_Err
            
148         If Not bNotAvailable Then

150             sConnectionStrings = sConnectionStrings & oRS.Fields("ConnectionString").value
152             iDDDefIndex = iDDDefIndex + 1
        
154             ReDim Preserve DDDefs(UBound(DDDefs) + 1)
156             DDDefs(UBound(DDDefs)).Name = oRS.Fields("DDDefName").value
158             DDDefs(UBound(DDDefs)).desc = oRS.Fields("Description").value
160             DDDefs(UBound(DDDefs)).Prefix = "dd_" & oRS.Fields("DDDefName").value & "_"
162             DDDefs(UBound(DDDefs)).ListIndex = iDDDefIndex - 1
164             DDDefs(UBound(DDDefs)).ConnectionString = oRS.Fields("ConnectionString").value

166             GetQueries oConn, DDDefs(UBound(DDDefs)), "dd_" & oRS.Fields("DDDefName").value & "_"

168             'If Not UBound(DDDefs(UBound(DDDefs)).Queries) = 0 Then
170                 cmbDatabases.AddItem oRS.Fields("Description").value, iDDDefIndex - 1
                'Else
172                 'iDDDefIndex = iDDDefIndex - 1
                'End If
        
            End If
        
174         If oConn.State = adStateOpen Then oConn.Close
176         oRS.MoveNext
        
        Loop
        
178     oRS.Close
180     Set oRS = Nothing

        Exit Sub
        
NotAvailable:
        bNotAvailable = True
        Resume Next
        '<EhFooter>Trim$
        Exit Sub

Init_Err:
        'MsgBox "DynamicReports.Init_Err (line " & Erl & "): " & Err.Description
        'Resume Next
        MsgBox "Your configuration settings for this module are in error.  Please contact an OASIS administrator", vbExclamation, "Configuration error"
        oRS.Close
        Set oRS = Nothing
        '</EhFooter>
        Exit Sub
        MsgBox "DynamicReports.Init_Err (line " & Erl & "): " & Err.Description
        Resume Next
End Sub

Private Function GetChartImage() As StdPicture
        '<EhHeader>
        On Error GoTo GetChartImage_Err
        '</EhHeader>
        Dim oPic As StdPicture
       
100     OASISChartingVer2.ImageSave g_sAppPath & "\data\user\exports\charttemp4.bmp"
102     Set oPic = LoadPicture(g_sAppPath & "\data\user\exports\charttemp4.bmp")
104     Set GetChartImage = oPic
        '<EhFooter>
        Exit Function

GetChartImage_Err:
    MsgBox "DynamicReports.GetChartImage_Err (line " & Erl & "): " & Err.Description
        
        '</EhFooter>
End Function

Private Sub SaveChartToClipboard()
        '<EhHeader>
        On Error GoTo SaveChartToClipboard_Err
        '</EhHeader>

        Dim oPic As StdPicture
        Clipboard.Clear
100     OASISChartingVer2.ImageSave g_sAppPath & "\data\user\exports\charttemp4.bmp"
102     Set oPic = LoadPicture(g_sAppPath & "\data\user\exports\charttemp4.bmp")
104     Clipboard.SetData oPic, vbCFDIB
106     Set oPic = Nothing
    
        '<EhFooter>
        Exit Sub

SaveChartToClipboard_Err:
 MsgBox "DynamicReports.SaveChartToClipboard_Err (line " & Erl & "): " & Err.Description
  
        '</EhFooter>
End Sub

Private Sub PointAtCurrentDDInList()
        '<EhHeader>
        On Error GoTo PointAtCurrentDDInList_Err
        '</EhHeader>

        Dim i As Long
100     i = 0
    
102     Do Until i = UBound(DDDefs) + 1
    
104         If DDDefs(i).ListIndex = cmbDatabases.ListIndex Then
106             DDDefCurrent = DDDefs(i)
            End If

108         i = i + 1
        Loop

        '<EhFooter>
        Exit Sub

PointAtCurrentDDInList_Err:
        MsgBox "DynamicReports.PointAtCurrentDDInList_Err (line " & Erl & "): " & Err.Description
        'Resume Next
        '</EhFooter>
End Sub


Public Sub cmbDatabases_Click()
        '<EhHeader>
        On Error GoTo cmbDatabases_Click_Err
        '</EhHeader>

        Dim lCountOfQueries As Long
        
        C1Tab1.Enabled = True
        dxDateEditFrom.Tag = "no event"
        dxDateEditFrom = Format(Now() - 14, "Medium Date")
        dxDateEditTill = Format(Now(), "Medium Date")
        dxDateEditFrom.Tag = ""
        
100     Call PointAtCurrentDDInList
        listGroups.Enabled = True
        listGroups.Clear
        listGroups.AddItem " -- ALL --"
        listGroups.ListIndex = 0
    
104     Set mConn = New ADODB.Connection
106     mConn.ConnectionString = DDDefCurrent.ConnectionString
108     mConn.ConnectionString = Replace(mConn.ConnectionString, "\data\db\dynamicdata", AppPath & "data\db\dynamicdata", , , vbTextCompare)
        If bSQLServerInUse Then mConn.ConnectionString = g_sGlobalConnectionString
110     mConn.CursorLocation = g_sGlobalCursorLocation

        On Error GoTo NoSuchDatabase
112     mConn.Open
        On Error GoTo cmbDatabases_Click_Err
    
        listQueries.Clear
        GetQueries mConn, DDDefCurrent, DDDefCurrent.Prefix
        GetGroups mConn, DDDefCurrent.Prefix
114     lCountOfQueries = UBound(DDDefCurrent.Queries)


        ShowQueries 'lCountOfQueries


        
        OASISChartingVer2.Visible = False
        C1ElasticQuery.Visible = True
        C1ElasticGrid.Visible = True
        C1Tab2.Visible = True
        Set dxDBGrid1.DataSource = Nothing
        dxDBGrid1.Columns.DestroyColumns
        txtTxtGroup.Text = ""
        dxFilterSQL = ""
     dxSQLCommand = ""
     chkAutoLoad.value = vbUnchecked
     chkUseChart.value = vbChecked
        
        Exit Sub

NoSuchDatabase:
136     MsgBox "This database is not available.", vbInformation
        'Resume Next
        '<EhFooter>
        Exit Sub

cmbDatabases_Click_Err:
        MsgBox "DynamicReports.cmbDatabases_Click_Err (line " & Erl & "): " & Err.Description
        'Resume Next
        '</EhFooter>
End Sub

Private Sub ShowQueries()

    Dim lQueryIndex As Long
    Dim sQueryName As String
    Dim i As Long

    i = 1
    listQueries.Clear
    lQueryIndex = 0

    If Len(DDDefCurrent.Name) > 1 Then

    Do Until i = UBound(DDDefCurrent.Queries) + 1
    
        With DDDefCurrent.Queries(i)
        
            .ListIndex = 666666
            sQueryName = .QueryName
            If (.Group = listGroups.Text) Or (listGroups.Text = " -- ALL --") Then listQueries.AddItem sQueryName ', lQueryIndex
            .ListIndex = lQueryIndex
            lQueryIndex = lQueryIndex + 1

        End With

        i = i + 1
    Loop
    
    End If
        
End Sub

Private Sub cmdDoubleChartSize_Click()
OASISChartingVer2.ResizeHeight 200
End Sub

Private Sub cmdExport_Click()
        '<EhHeader>
        On Error GoTo cmdExport_Click_Err
        '</EhHeader>

        Dim c As New cCommonDialog
        Dim objPic As StdPicture
    
        Dim m_frmReportsFromRS As New frmReportsFromRS
        
102     If Not mRS Is Nothing Then
   
104         Select Case lstExports.Text
    
                Case "Data to Excel"
            
106                 With c
108                     .CancelError = False
110                     .DialogTitle = "Export data to..."
112                     .InitDir = g_sAppPath & "\data\"
114                     .Filter = "Microsoft Excel (.xls)|*.xls"
116                     .DefaultExt = ".xls"
118                     .ShowSave
                    End With
            
120                 If c.ConfirmOK And Not c.Filename = "" And Not IsNull(c.Filename) Then
                        CreateXLSFromRS mRS, c.Filename
122                     'dxDBGrid1.M.ExportToXLS c.Filename
                    End If

124             Case "Data to HTML"
            
126                 With c
128                     .CancelError = False
130                     .DialogTitle = "Export data to..."
132                     .InitDir = g_sAppPath & "\data\"
134                     .Filter = "HTML Standard Format (.html)|*.html"
136                     .DefaultExt = ".html"
138                     .ShowSave
                    End With
            
140                 If c.ConfirmOK And Not c.Filename = "" And Not IsNull(c.Filename) Then
142                     dxDBGrid1.M.ExportToHTML c.Filename
                    End If

144             Case "Data to XML"

146                 With c
148                     .CancelError = False
150                     .DialogTitle = "Export data to..."
152                     .InitDir = g_sAppPath & "\data\"
154                     .Filter = "XML Standard Format (.xml)|*.xml"
156                     .DefaultExt = ".xml"
158                     .ShowSave
                
                    End With
                
160                 If c.ConfirmOK And Not c.Filename = "" And Not IsNull(c.Filename) Then
162                     dxDBGrid1.M.ExportToXML c.Filename
                    End If

164             Case "Data to OASIS Reports"

166                 m_frmReportsFromRS.SetReportRS cmbDatabases.Text & " [" & listQueries.Text & "]", mRS, ""
168                 m_frmReportsFromRS.ShowReport
170                 m_frmReportsFromRS.Show vbModal, Me

172             Case "Chart to Image File"
 
174                 With c
176                     .CancelError = False
178                     .DialogTitle = "Export chart to..."
180                     .InitDir = g_sAppPath & "\data\"
182                     .Filter = "Bitmap Image (.bmp)|*.bmp"
184                     .DefaultExt = ".bmp"
186                     .ShowSave
                    End With
                
188                 If c.ConfirmOK And Not c.Filename = "" And Not IsNull(c.Filename) Then
            
                        OASISChartingVer2.ImageSave c.Filename
          
                    End If

194             Case "Chart to Clipboard"
        
196                 SaveChartToClipboard
                    MsgBox "The chart has been copied to the clipboard"
                    

198             Case "Chart and Data to OASIS Reports"
               
200                 Set objPic = GetChartImage
202                 m_frmReportsFromRS.SetReportRS cmbDatabases.Text & " [" & listQueries.Text & "]", mRS, "", objPic
204                 m_frmReportsFromRS.ShowReport
206                 m_frmReportsFromRS.Show vbModal, Me
    
            End Select
    
        End If
    
208     Set m_frmReportsFromRS = Nothing
210     Set c = Nothing
212     Set objPic = Nothing

        '<EhFooter>
        Exit Sub

cmdExport_Click_Err:
        MsgBox "DynamicReports.cmdExport_Click_Err (line " & Erl & "): " & Err.Description
        'Resume Next
        '</EhFooter>
End Sub

Private Sub cmdHalfChatrSize_Click()
    OASISChartingVer2.ResizeHeight 50
End Sub

Private Sub cmdMakeChart_Click()
OASISChartingVer2.MakeVisible
End Sub

Private Sub cmdMakeChartIn_Click()
OASISChartingVer2.MakeInvisible

End Sub

Private Sub cmdRefreshChart_Click()
    OASISChartingVer2.RefreshChart
End Sub

Private Sub cmdRemoveChart_Click()
    RemoveChartFromDB
End Sub

Private Sub cmdRepaint_Click()
  ' OASISChartingVer2.up
   C1ElasticChart.Refresh
     UserControl.Refresh
    
End Sub

Private Sub cmdREsetChart_Click()
    OASISChartingVer2.ResetChart
End Sub

Private Sub cmdSaveSQL_Click()

    If UpdateReport Then SaveChartSettings
End Sub

Private Function UpdateReport() As Boolean
        '<EhHeader>
        On Error GoTo UpdateReport_Err
        '</EhHeader>

        Dim oRS As ADODB.Recordset
        Dim mstream As New ADODB.Stream

100     UpdateReport = False
        
        On Error GoTo SyntaxSQLError
102     Set oRS = New ADODB.Recordset
104     oRS.Open CStr(dxSQLCommand), mConn, adOpenStatic, adLockReadOnly

106     If oRS.State = adStateOpen Then oRS.Close
108     Set oRS = Nothing

        If Len(CStr(dxFilterSQL)) > 0 Then
            On Error GoTo SyntaxFilterSQLError
110         Set oRS = New ADODB.Recordset
112         oRS.Open CStr(dxFilterSQL), mConn, adOpenStatic, adLockReadOnly
    
114         If oRS.State = adStateOpen Then oRS.Close
116         Set oRS = Nothing
        End If

118     UpdateReport = True
        On Error GoTo UpdateReport_Err
        'OASISChartingVer2.ClearDataAll
119     OASISChartingVer2.TemplateSave g_sAppPath & "\data\user\exports\charttempo.oct"
120     DDQueryCurrent.AutoLoadChart = IIf(chkAutoLoad.value = vbChecked, True, False)
122     DDQueryCurrent.Command = dxSQLCommand
124     DDQueryCurrent.FilterSQL = dxFilterSQL
126     DDQueryCurrent.UseChart = IIf(chkUseChart.value = vbChecked, True, False)
    
128     If Not DDQueryCurrent.Group = txtTxtGroup.Text Then
130         MsgBox "Note you have to save changes and reload the reporting module to reset the grouping.", vbInformation
        End If
    
132     DDQueryCurrent.Group = txtTxtGroup.Text
    
        'On Error Resume Next
134     mstream.Type = adTypeBinary
136     mstream.Open
138     mstream.LoadFromFile g_sAppPath & "\data\user\exports\charttempo.oct"
140     DDQueryCurrent.ChartSettings = mstream.Read
142     mstream.Close
144     Set mstream = Nothing

        If UBound(DDDefCurrent.Queries) > 0 Then

146         With DDQueryCurrent
148             DDDefCurrent.Queries(lCurrentQueryIndex).AutoLoadChart = .AutoLoadChart
150             DDDefCurrent.Queries(lCurrentQueryIndex).Command = .Command
152             DDDefCurrent.Queries(lCurrentQueryIndex).FilterSQL = .FilterSQL
154             DDDefCurrent.Queries(lCurrentQueryIndex).Group = .Group
156             DDDefCurrent.Queries(lCurrentQueryIndex).UseChart = .UseChart
158             DDDefCurrent.Queries(lCurrentQueryIndex).ChartSettings = .ChartSettings
            End With
 
160         LoadFromQueriesList False
        End If

        '<EhFooter>
        Exit Function

SyntaxSQLError:
        MsgBox "The SQL Command syntax is invalid: " & Err.Description

        If oRS.State = adStateOpen Then oRS.Close
        Set oRS = Nothing
        Exit Function
        
SyntaxFilterSQLError:
        MsgBox "The SQL Filter syntax is invalid: " & Err.Description

        If oRS.State = adStateOpen Then oRS.Close
        Set oRS = Nothing
        Exit Function
        
UpdateReport_Err:
        MsgBox "DynamicReports.UpdateReport_Err (line " & Erl & "): " & Err.Description
        UpdateReport = False
        '</EhFooter>
End Function

Private Sub cmdUpdateReport_Click()
    Call UpdateReport
End Sub


Private Sub dxDateEditFrom_Change()
 If Not dxDateEditFrom.Tag = "no event" Then DisplayChartSelected False
'Call dxDateEditTill_Change
End Sub

Private Sub dxDateEditTill_Change()
    
 If Not dxDateEditFrom.Tag = "no event" Then DisplayChartSelected False
    
'    Dim sSQL As String
'
'    sSQL = dxSQLCommand
'
'    If dxDateEditFrom.Visible And Len(dxDateEditFrom) > 0 Then sSQL = Replace(sSQL, "#01-01-1900#", "#" & dxDateEditFrom & "#")
'    If dxDateEditTill.Visible And Len(dxDateEditTill) > 0 Then sSQL = Replace(sSQL, "#01-01-9999#", "#" & dxDateEditTill & "#")
'
'    Set mRS = New ADODB.Recordset
'    If Len(sSQL) > 5 Then
'    mRS.Open sSQL, mConn, adOpenDynamic, adLockBatchOptimistic
'    OASISChartingVer2.LoadRS mRS
'    End If
'
'    Set dxDBGrid1.DataSource = mRS
'    If Not mRS.State = 0 Then C1ERecordCount.caption = "Record Count: " & mRS.RecordCount
    
End Sub


Private Sub dxDBGrid1_OnFilterChanged()

    OASISChartingVer2.LoadRS mRS
End Sub

Private Sub dxDBGrid1_OnHeaderClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn)

    On Error Resume Next

    If mRS.Sort = "[" & Column.FieldName & "]" Then
        mRS.Sort = "[" & Column.FieldName & "] DESC"
    Else
        mRS.Sort = "[" & Column.FieldName & "]"
    End If
    
    OASISChartingVer2.LoadRS mRS

End Sub


' old function which listed all queries with prefix
'
'Private Sub GetQueries(oCN As ADODB.Connection, _
'                       DynamicDataDef As DYNAMIC_DATA_DEF, _
'                       sPreFix As String)
'        '<EhHeader>
'        On Error GoTo GetQueries_Err
'        '</EhHeader>
'
'        Dim i As Long
'        Dim cat As ADOx.Catalog
'        Dim cmd As ADODB.Command
'100     Set cat = New ADOx.Catalog
'102     Set cat.ActiveConnection = oCN
'
'104     ReDim Preserve DynamicDataDef.Queries(0)
'
'106     i = 0
'
'108     Do Until i = cat.Views.Count
'
'110         If Left$(cat.Views(i).Name, Len(sPreFix)) = sPreFix And Not Right$(cat.Views(i).Name, 4) = "_FEA" And Not Right$(cat.Views(i).Name, 4) = "_GEO" And Not Right$(cat.Views(i).Name, 5) = "_HIDE" Then
'
'112             ReDim Preserve DynamicDataDef.Queries(UBound(DynamicDataDef.Queries) + 1)
'
'114             With DynamicDataDef.Queries(UBound(DynamicDataDef.Queries))
'
'116                 .QueryName = Right(cat.Views(i).Name, Len(cat.Views(i).Name) - Len(sPreFix))
'118                 Set cmd = cat.Views(i).Command
'120                 .Command = cmd.CommandText
'
'                End With
'
'            End If
'
'122         i = i + 1
'        Loop
'
'124     Set cmd = Nothing
'126     Set cat = Nothing
'
'        '<EhFooter>
'        Exit Sub
'
'GetQueries_Err:
'        MsgBox "DynamicReports.GetQueries_Err (line " & Erl & "): " & Err.Description
'        'Resume Next
'        '</EhFooter>
'End Sub

Private Sub GetGroups(oCn As ADODB.Connection, _
                      sPreFix As String)
        
    Dim oRS As New ADODB.Recordset
    Dim sSQL As String
        
    sSQL = "SELECT DISTINCT [Group] FROM [" & sPreFix & "ChartSettings] ORDER BY [Group]"
    oRS.Open sSQL, oCn, adOpenDynamic, adLockReadOnly
        
    If Not oRS Is Nothing Then
        
        If Not oRS.State = adStateClosed Then
            
            Do Until oRS.EOF
          
                If (Len(oRS.Fields(0).value) > 0) Then listGroups.AddItem oRS.Fields(0).value
                oRS.MoveNext
            
            Loop
            
            oRS.Close
            
        End If
        
    End If

    Set oRS = Nothing

End Sub

Private Sub GetQueries(oCn As ADODB.Connection, _
                       DynamicDataDef As DYNAMIC_DATA_DEF, _
                       sPreFix As String)
        '<EhHeader>
        On Error GoTo GetQueries_Err
        '</EhHeader>
        
        Dim oRS As New ADODB.Recordset
        Dim sSQL As String
        'Dim oCN As New ADODB.Connection
        
100     ReDim DynamicDataDef.Queries(0)

102     If DoesTableExist(IIf(bSQLServerInUse, g_sGlobalConnectionString, oCn.ConnectionString), sPreFix & "ChartSettings") Then
        
104         sSQL = "SELECT * FROM [" & sPreFix & "ChartSettings] ORDER BY [QueryName]"
106         oRS.Open sSQL, oCn, adOpenDynamic, adLockReadOnly
        
108         If Not oRS Is Nothing Then
        
110             If Not oRS.State = adStateClosed Then
            
112                 Do Until oRS.EOF
          
114                     ReDim Preserve DynamicDataDef.Queries(UBound(DynamicDataDef.Queries) + 1)

116                     With DynamicDataDef.Queries(UBound(DynamicDataDef.Queries))
    
118                         .QueryName = oRS.Fields("QueryName").value

120                         If bSQLServerInUse Then
122                             .Command = IIf(IsNull(oRS.Fields("MSSQLCommand").value), "", oRS.Fields("MSSQLCommand").value)
124                             .FilterSQL = IIf(IsNull(oRS.Fields("FilterMSSQL").value), "", oRS.Fields("FilterMSSQL").value)
                            Else
126                             .Command = IIf(IsNull(oRS.Fields("SQLCommand").value), "", oRS.Fields("SQLCommand").value)
128                             .FilterSQL = IIf(IsNull(oRS.Fields("FilterSQL").value), "", oRS.Fields("FilterSQL").value)
                            End If

130                         .UseChart = oRS.Fields("UseChart").value
132                         .Group = IIf(IsNull(oRS.Fields("Group").value), "", oRS.Fields("Group").value)
134                         .AutoLoadChart = oRS.Fields("bAutoLoadReport").value
136                         If Not IsNull(oRS.Fields("OCTSettings").value) Then .ChartSettings = oRS.Fields("OCTSettings").value
                            
                        End With
            
138                     oRS.MoveNext
            
                    Loop
            
140                 oRS.Close
            
                End If
        
            End If
        End If

142     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

GetQueries_Err:
        Set oRS = Nothing
        MsgBox "DynamicReports.GetQueries_Err (line " & Erl & "): " & Err.Description
        
'        Stop
 '       Resume Next
        
        '        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadFromQueriesList(bLoadFromFile As Boolean)
        '<EhHeader>
        On Error GoTo LoadFromQueriesList_Err
        '</EhHeader>
        Dim oRS As ADODB.Recordset

100     CheckForDateRange

102     If IsNull(dxDateEditFrom) Then dxDateEditFrom = Format(Now() - 14, "Medium Date")
104     If IsNull(dxDateEditTill) Then dxDateEditTill = Format(Now(), "Medium Date")
106     OASISChartingVer2.Visible = True
108     OASISChartingVer2.Width = OASISChartingVer2.Width * 2
110     C1ElasticChart.Refresh
112     OASISChartingVer2.RefreshChart
        'OASISChartingVer2.ClearLegend
114     cmbFilter.Clear
116     cmbFilter.Enabled = False
        
118     If Not IsNull(listQueries.Text) And Not listQueries.Text = "" Then
                            
            OASISChartingVer2.ClearDataLabels
120         Call PointAtCurrentQueryInList

122         If Len(DDQueryCurrent.FilterSQL) > 1 Then
            
124             Set oRS = New ADODB.Recordset
126             oRS.Open DDQueryCurrent.FilterSQL, mConn, adOpenDynamic, adLockReadOnly
                
128             If Not oRS.State = 0 Then
                
130                 cmbFilter.AddItem "  -- ALL --"
                
132                 Do Until oRS.EOF
                    
134                     cmbFilter.Enabled = True
136                     cmbFilter.AddItem oRS.Fields(0).value
138                     oRS.MoveNext
                        
                    Loop
                    
140                 If cmbFilter.ListCount > 0 Then
142                     cmbFilter.Tag = "no event"
144                     cmbFilter.ListIndex = 0
146                     cmbFilter.Tag = ""
148                     DisplayChartSelected bLoadFromFile
                    End If
                End If
            
            Else
            

150             DisplayChartSelected bLoadFromFile
            
            End If

        End If

152     CheckForDateRange

        '<EhFooter>
        Exit Sub

LoadFromQueriesList_Err:
        MsgBox "It looks like the query has an error in it - please contact an OASIS Administrator and quote the following error: " & Chr(13) & Chr(13) & "(" & Erl & ")" & Err.Description
        
        '</EhFooter>
End Sub

Public Sub listQueries_Click()
         LoadFromQueriesList True
End Sub

Public Sub DisplayChartSelected(bLoadFromFile As Boolean)
        '<EhHeader>
        On Error GoTo DisplayChartSelected_Err
        '</EhHeader>

        Dim sSQL As String
        Dim bLoadingSQL As Boolean
100     C1ElasticChart.caption = "Loading: Please wait a moment................"
102     OASISChartingVer2.Visible = False
104     C1Tab2.Visible = False

106     DoEvents
    
108     sSQL = DDQueryCurrent.Command
    
110     If Len(sSQL) > 0 Then
    
            Dim i As Long
            Dim lColWidths As Long
            Dim sFilter As String
            Dim oRS As ADODB.Recordset
            Dim bBand1 As dxGridBand
            Dim bBand2 As dxGridBand
            Dim sumgroup  As dxGridSummaryGroup
            Dim Col As dxGridColumn
        
112         Set mRS = Nothing
114         dxSQLCommand = ""
116         LoadChartTemplate bLoadFromFile
118         dxSQLCommand = sSQL
120         txtTxtGroup.Text = DDQueryCurrent.Group
122         dxFilterSQL = DDQueryCurrent.FilterSQL
124         txtTxtGroup = DDQueryCurrent.Group
126         CheckForDateRange

            If bSQLServerInUse Then

                ' If dxDateEditFrom.Visible And Len(dxDateEditFrom) > 0 Then
                sSQL = Replace(sSQL, "'01-01-1900'", "'" & dxDateEditFrom & "'")
                ' If dxDateEditTill.Visible And Len(dxDateEditTill) > 0 Then
                sSQL = Replace(sSQL, "'01-01-9999'", "'" & dxDateEditTill & "'")

            Else
128             '   If dxDateEditFrom.Visible And Len(dxDateEditFrom) > 0 Then
                sSQL = Replace(sSQL, "#01-01-1900#", "#" & dxDateEditFrom & "#")
130             '   If dxDateEditTill.Visible And Len(dxDateEditTill) > 0 Then
                sSQL = Replace(sSQL, "#01-01-9999#", "#" & dxDateEditTill & "#")

            End If

132         If Len(cmbFilter.Text) > 0 Then sSQL = Replace(sSQL, "XxXxXx", Replace(cmbFilter.Text, "'", "''"))
134         sSQL = Replace(sSQL, "= '  -- ALL --'", "<> '  -- ALL --'")
136         sSQL = Replace(sSQL, "='  -- ALL --'", "<> '  -- ALL --'")
        
138         Set mRS = New ADODB.Recordset
140         bLoadingSQL = True
142         mRS.Open sSQL, mConn, adOpenDynamic, adLockBatchOptimistic
144         bLoadingSQL = False
        
146         If mRS.State = adStateOpen Then
        
148             If mRS.RecordCount > 1000 Then
150                 dxDBGrid1.Option = egoLoadAllRecords
152                 dxDBGrid1.OptionEnabled = False
        
154                 dxDBGrid1.Option = egoDynamicLoad
156                 dxDBGrid1.OptionEnabled = True
        
                Else
        
158                 dxDBGrid1.Option = egoLoadAllRecords
160                 dxDBGrid1.OptionEnabled = True
        
162                 dxDBGrid1.Option = egoDynamicLoad
164                 dxDBGrid1.OptionEnabled = False
                End If

166             i = dxDBGrid1.Bands.Count - 1

168             Do Until i < 0
170                 dxDBGrid1.Bands(i).Delete
172                 i = i - 1
                Loop
                
174             For i = 0 To dxDBGrid1.SummaryGroups.Count
176                 dxDBGrid1.SummaryGroups.Remove i
                Next
                
                'dxDBGrid1.Option = egoShowFooter
                'dxDBGrid1.OptionEnabled = True
178             dxDBGrid1.Columns.DestroyColumns
                
180             Set bBand1 = dxDBGrid1.Bands.Add
182             Set bBand2 = dxDBGrid1.Bands.Add
184             bBand1.ObjectName = mRS.Fields(0).Name
186             bBand1.caption = mRS.Fields(0).Name
188             bBand1.Fixed = bfLeft
190             bBand1.MinWidth = C1ElasticQuery.Width / 8.2 / Screen.TwipsPerPixelX
192             bBand1.Width = C1ElasticQuery.Width / 8.2 / Screen.TwipsPerPixelX

194             bBand2.ObjectName = "data"
196             bBand2.caption = "Data"
198             bBand1.Fixed = bfLeft
200             bBand1.Visible = True
202             bBand2.Visible = True
204             dxDBGrid1.ScrollBars = ssBoth
    
206             dxDBGrid1.Option = egoShowBands
208             dxDBGrid1.OptionEnabled = False
210             dxDBGrid1.Option = egoAutoWidth
212             dxDBGrid1.OptionEnabled = False
    
214             dxDBGrid1.KeyField = mRS.Fields(0).Name
216             Set dxDBGrid1.DataSource = mRS
218             dxDBGrid1.Columns.RetrieveFields
220             lColWidths = C1ElasticQuery.Width / 8.2 / Screen.TwipsPerPixelX
            
222             i = 0

224             Do Until i = dxDBGrid1.Columns.Count
226                 dxDBGrid1.Columns(i).Width = lColWidths
                    
228                 If i = 0 Then
230                     dxDBGrid1.Columns(i).bandIndex = bBand1.Index
                    Else
232                     dxDBGrid1.Columns(i).bandIndex = bBand2.Index
                    End If
                    
234                 i = i + 1
                Loop
                
236             Set sumgroup = dxDBGrid1.SummaryGroups.Add
238             sumgroup.DefaultGroup = True
                'Set sumitem = sumgroup.SummaryItems.Add
            
                'dxDBGrid1.Columns(0).SummaryFooterFormat = "0 records"
                'sumitem.SummaryType = cstCount

                ' For i = 0 To dxDBGrid1.Columns.Count - 1
                '   Set col = dxDBGrid1.Columns(i)

                'If col.Visible Then
                '   col.SummaryFooterType = cstCount
                '  Exit For
                'End If

                'Next
            
290             C1ERecordCount.caption = "Record Count: " & mRS.RecordCount

If DDQueryCurrent.UseChart Then OASISChartingVer2.LoadRS dxDBGrid1.DataSource  ' mRS
                
292             If DDQueryCurrent.AutoLoadChart Then
                
294                 If DDQueryCurrent.UseChart Then
296                     lstExports.Text = "Chart and Data to OASIS Reports"
                    Else
298                     lstExports.Text = "Data to OASIS Reports"
                    End If

300                 Call cmdExport_Click
                
                End If
                
                
            End If

        End If
        
240     If DDQueryCurrent.AutoLoadChart Then
242         chkAutoLoad.value = vbChecked
        Else
244         chkAutoLoad.value = vbUnchecked
        End If
    
246     If DDQueryCurrent.UseChart Then
                    
248         chkUseChart.value = vbChecked
250         C1Tab1.Visible = False
252         C1Tab2.Visible = False
254         C1Tab1.top = 0
256         C1Tab1.Height = C1Elastic1.Height * 0.75
258         C1Tab2.top = C1Elastic1.Height - 12
260         C1Tab2.Height = C1Elastic1.Height / 4
                    
262         C1Elastic1.GridCols = 1
264         C1Elastic1.GridRows = 3
266         C1Tab1.Visible = True
268         C1Elastic1.Refresh

270         DoEvents
                        
272         ' OASISChartingVer2.LoadRS dxDBGrid1.DataSource  ' mRS

        Else
274         chkUseChart.value = vbUnchecked
276         C1Tab1.Visible = False
278         C1Tab2.Visible = True
280         C1Elastic1.GridCols = 1
282         C1Elastic1.GridRows = 1

284         DoEvents
        End If

        '  C1Elastic1.Refresh
286     C1Tab2.CurrTab = 0
        'DoEvents
288     C1Tab2.Visible = True

302     C1Tab2.Visible = True
304     OASISChartingVer2.Visible = True
306     C1ElasticChart.caption = ""
    
        '<EhFooter>
        Exit Sub

DisplayChartSelected_Err:

        If bLoadingSQL Then
            MsgBox "There is an error in the SQL statement! (" & Err.Description & ")"
            Resume Next
        Else
            MsgBox "DynamicReports.DisplayChartSelected_Err (line " & Erl & "): " & Err.Description
        End If

        '</EhFooter>
End Sub


Private Sub CheckForDateRange()

    If InStr(dxSQLCommand, "#01-01-1900#") > 0 Or InStr(dxSQLCommand, "'01-01-1900'") > 0 Then
    
        C1EDateFrom.Visible = True
        dxDateEditFrom.Visible = True
        
    
    Else
    
        C1EDateFrom.Visible = False
        dxDateEditFrom.Visible = False
    
    End If
    
    If InStr(dxSQLCommand, "#01-01-9999#") > 0 Or InStr(dxSQLCommand, "'01-01-9999'") > 0 Then
    
        C1EDateTill.Visible = True
        dxDateEditTill.Visible = True
        
    
    Else
    
        C1EDateTill.Visible = False
        dxDateEditTill.Visible = False
    
    End If

End Sub


Private Sub PointAtCurrentQueryInList()
        '<EhHeader>
        On Error GoTo PointAtCurrentQueryInList_Err
        '</EhHeader>

        Dim i As Long
100     i = 0
        lCurrentQueryIndex = -1
    
102     Do Until i = UBound(DDDefCurrent.Queries) + 1
    
104         If DDDefCurrent.Queries(i).QueryName = listQueries.Text Then
106             DDQueryCurrent = DDDefCurrent.Queries(i)
                lCurrentQueryIndex = i
            End If

108         i = i + 1
        Loop

        '<EhFooter>
        Exit Sub

PointAtCurrentQueryInList_Err:
        MsgBox "DynamicReports.PointAtCurrentQueryInList_Err (line " & Erl & "): " & Err.Description
        'Resume Next
        '</EhFooter>
End Sub

Private Function FileExists(sFile As String) As Integer
        '<EhHeader>
        On Error GoTo FileExists_Err
        '</EhHeader>
        On Error Resume Next

100     FileExists = (Dir$(sFile) <> "")

        '<EhFooter>
        Exit Function

FileExists_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.FileExists " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub SaveChartSettings()
        '<EhHeader>
        On Error GoTo SaveChartSettings_Err
        '</EhHeader>
    
        Dim mstream As New ADODB.Stream
        Dim oRS As New ADODB.Recordset
        Dim sGUID As String
        Dim bEdit As Boolean
        Dim bAddNew As Boolean
        Dim sNewName As String
        Dim sQueryName As String
    
        'Call PointAtCurrentDDInList
        'Call PointAtCurrentQueryInList
        
100     'OASISChartingVer2.TemplateSave g_sAppPath & "\data\user\exports\charttempo.oct"
102     'OASISChartingVer2.RefreshChart
        
104     If Not listQueries.Text = "" And Not IsNull(listQueries.Text) And Len(listQueries.Text) > 0 Then
106         bAddNew = IIf(MsgBox("Do you want to save to existing report [" & listQueries.Text & "]?  Clicking 'No' will prompt you for a new report name", vbYesNo, "Add new or edit") = vbYes, False, True)
        Else
108         bAddNew = True
        End If
        
110     sNewName = ""

112     If bAddNew Then
        
114         sNewName = InputBox("Please enter a new report name:", "New report name", "")

116         If Not Len(sNewName) > 0 Then
            
118             MsgBox "No report name specified! Aborting save.", vbCritical
                Exit Sub
            End If
            
        End If
        
       ' OASISChartingVer2.TemplateSave g_sAppPath & "\data\user\exports\charttempo.oct"
            
120     If bAddNew Then
122         sQueryName = sNewName
        Else
124         sQueryName = DDQueryCurrent.QueryName
        End If
        
126     With oRS

128         .Open "SELECT * FROM [" & DDDefCurrent.Prefix & "ChartSettings] WHERE QueryName = '" & sQueryName & "'", mConn, adOpenDynamic, adLockBatchOptimistic

130         If .State = adStateOpen Then
    
132             If .EOF Then
                    
134                 .AddNew
136                 .Fields("GUID1").value = GUIDGen
138                 .Fields("QueryName").value = sQueryName
                    
                Else
                
140                 If bAddNew Then
                    
142                     MsgBox "This report name already exists! Aborting save.", vbCritical
144                     .Close
                        Exit Sub
                        
                    End If
                    
146                 bEdit = True
                End If
            
148             sGUID = .Fields("GUID1").value
                
150             If bSQLServerInUse Then
152                 .Fields("MSSQLCommand").value = dxSQLCommand
154                 .Fields("FilterMSSQL").value = DDQueryCurrent.FilterSQL
                Else
156                 .Fields("SQLCommand").value = dxSQLCommand
158                 .Fields("FilterSQL").value = DDQueryCurrent.FilterSQL
                End If

                If FileExists(g_sAppPath & "\data\user\exports\charttempo.oct") Then
                    mstream.Type = adTypeBinary
                    mstream.Open
                    mstream.LoadFromFile g_sAppPath & "\data\user\exports\charttempo.oct"
160                 .Fields("OCTSettings").value = mstream.Read
                    mstream.Close
                End If

                Set mstream = Nothing
                
162             .Fields("UseChart").value = DDQueryCurrent.UseChart
164             .Fields("bAutoLoadReport").value = DDQueryCurrent.AutoLoadChart
166             .Fields("Group").value = DDQueryCurrent.Group
168             .UpdateBatch adAffectCurrent
170             .Close
            
172             If Not bEdit Then
174                 SynchHistoryAddNew GUIDGen, sGUID, "Chart Synch", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & "ChartSettings", False, "false", IIf(DDDefCurrent.Prefix = "Incidents_", "", DDDefCurrent.Prefix)
176                 GetQueries mConn, DDDefCurrent, DDDefCurrent.Prefix
178                 listQueries.AddItem sQueryName
                Else
180                 SynchHistoryEdit GUIDGen, sGUID, "Chart Synch", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & "ChartSettings", False, "false", IIf(DDDefCurrent.Prefix = "Incidents_", "", DDDefCurrent.Prefix)
                    'SynchHistoryEdit GUIDGen, mDDRSLinked.Fields(1).Value, "DD linkedtable Edit", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & DDTableCurrent.TableName, False, "false", DDDefCurrent.Prefix
                End If
        
            End If
    
        End With
    
182     MsgBox "Save complete", vbInformation
       ' UpdateReport
        '<EhFooter>
        Exit Sub

SaveChartSettings_Err:
        MsgBox "Save Failed!..... DynamicReports.SaveChartSettings_Err (line " & Erl & "): " & Err.Description
        
        '</EhFooter>
End Sub

Private Sub RemoveChartFromDB()
        '<EhHeader>
        On Error GoTo RemoveChartFromDB_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
        Dim sGUID As String
        'OASISChartingVer2.ResetChart
        'listQueries.ListIndex = listQueries.lis

100     With oRS

102         .Open "SELECT * FROM [" & DDDefCurrent.Prefix & "ChartSettings] WHERE QueryName = '" & DDQueryCurrent.QueryName & "'", mConn, adOpenDynamic, adLockBatchOptimistic

104         If .State = adStateOpen Then
    
106             If Not .EOF Then
               
108                 sGUID = .Fields("GUID1").value
110                 .Delete adAffectCurrent
112                 .UpdateBatch
114                 SynchHistoryDelete GUIDGen, sGUID, "Chart Synch", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & "ChartSettings", False, "false", DDDefCurrent.Prefix
116                 listQueries.RemoveItem listQueries.ListIndex
118                 MsgBox "Deletion successful"
                
120                 If listQueries.ListCount > 0 Then
122                     listQueries.ListIndex = 0
124                     LoadFromQueriesList False
                    Else
126                     Call cmbDatabases_Click
                    End If
                
                Else
            
128                 MsgBox "Deletion unsuccessful: Record was not found in the database"
            
                End If
        
            Else
            
130             MsgBox "Deletion unsuccessful: The table did not open"
            
            End If
    
        End With
        
        '<EhFooter>
        Exit Sub

RemoveChartFromDB_Err:
        MsgBox "DynamicReports.RemoveChartFromDB_Err (line " & Erl & "): " & Err.Description
        
        '</EhFooter>
End Sub

Private Sub LoadChartTemplate(bLoadFromFile As Boolean)

    Dim mstream As New ADODB.Stream
    Dim oRS As New ADODB.Recordset

    If bLoadFromFile And Not mConn Is Nothing Then

        With oRS

            .Open "SELECT * FROM [" & DDDefCurrent.Prefix & "ChartSettings] WHERE QueryName = '" & DDQueryCurrent.QueryName & "'", mConn, adOpenDynamic, adLockBatchOptimistic

            If .State = adStateOpen Then

                If Not .EOF Then

                    'DDQueryCurrent.UseChart = .Fields("UseChart").Value

                    '  If Not Len(.Fields("OCTSettings").Value < 10) Then
                    If FileExists(g_sAppPath & "\data\user\exports\charttempo.oct") And Not IsNull(.Fields("OCTSettings").value) Then
                        mstream.Type = adTypeBinary
                        mstream.Open
                        mstream.Write .Fields("OCTSettings").value
                        mstream.SaveToFile g_sAppPath & "\data\user\exports\charttempo.oct", adSaveCreateOverWrite
                        mstream.Close
                    End If

                    '  End If

                    'DDQueryCurrent.Command = .Fields("SQLCommand").Value
                Else
                
                    'OASISChartingVer2.ResetChart
                    
                End If

                .Close
            Else
                
                'OASISChartingVer2.ResetChart
                
            End If

        End With
        
    Else
                
        'OASISChartingVer2.ResetChart
                    
    End If

    If FileExists(g_sAppPath & "\data\user\exports\charttempo.oct") Then
        OASISChartingVer2.TemplateLoad (g_sAppPath & "\data\user\exports\charttempo.oct")
    End If

    Set oRS = Nothing
End Sub



Private Sub UserControl_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF11 Then
    
        If cmdSaveSQL.Visible = True Then
            cmdSaveSQL.Visible = False
            cmdRemoveChart.Visible = False
            
        Else
            cmdSaveSQL.Visible = True
            cmdRemoveChart.Visible = True
        End If
        
    End If

End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    Set mRS = Nothing
    Set mConn = Nothing
End Sub

Private Function SynchHistoryAddNew(sNewGUID As String, _
                                    sID As String, _
                                    sTitle As String, _
                                    sDescription As String, _
                                    sBy As String, _
                                    sRFC3339DateTime As String, _
                                    sTableName As String, _
                                    bIsGeoLayer As Boolean, _
                                    supdates As String, _
                                    Optional sSynchHistPrefix As String = "")
        '<EhHeader>
        On Error GoTo SynchHistoryAddNew_Err
        '</EhHeader>

100     SynchHistoryAddNew = False

        Dim oRS As New ADODB.Recordset
102     oRS.Open "SELECT * FROM [" & sSynchHistPrefix & "SynchHistory] WHERE sID = '" & sID & "' AND sTableName = '" & sTableName & "'", mConn, adOpenDynamic, adLockBatchOptimistic

104     With oRS
 
106         If .State = adStateOpen Then
    
108             If .RecordCount > 0 Then
                    
110                 DebugPrint "(frmDynamicData.SynchHistoryAddNew) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be 0 record with sID: " & sID & " for tablename [" & sTableName & "] but there are actually " & .RecordCount & "!!!  These records will be deleted.  If this error message persists please contact an OASIS Developer"

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
            
148             DebugPrint "(frmDynamicData.SynchHistoryAddNew) Table [" & sSynchHistPrefix & "SynchHistory] failed to open"
            
            End If
        
        End With
        
150     Set oRS = Nothing
        SynchHistoryAddNew = True
        '<EhFooter>
        Exit Function

SynchHistoryAddNew_Err:
        SynchHistoryAddNew = False
        Set oRS = Nothing
        MsgBox "DynamicReports.SynchHistoryAddNew_Err (line " & Erl & "): " & Err.Description
        'Resume Next
        '</EhFooter>
End Function

Private Function SynchHistoryEdit(sNewGUID As String, _
                                  sID As String, _
                                  sTitle As String, _
                                  sDescription As String, _
                                  sBy As String, _
                                  sRFC3339DateTime As String, _
                                  sTableName As String, _
                                  bIsGeoLayer As Boolean, _
                                  supdates As String, _
                                  Optional sSynchHistPrefix As String = "", _
                                  Optional lSequence As Long = 0)
        '<EhHeader>
        On Error GoTo SynchHistoryEdit_Err
        '</EhHeader>

100     SynchHistoryEdit = False
    
        Dim oRS As New ADODB.Recordset
102     oRS.Open "SELECT * FROM [" & sSynchHistPrefix & "SynchHistory] WHERE sID = '" & sID & "' AND sTableName = '" & sTableName & "' ORDER by [sequence] DESC", mConn, adOpenDynamic, adLockBatchOptimistic

104     With oRS
 
106         If .State = adStateOpen Then
    
108             If Not .EOF Then

                    '.Fields("sID").Value = sID
110                 .Fields("sGUID").value = sNewGUID
112                 .Fields("sTableName").value = sTableName
114                 .Fields("swhen").value = sRFC3339DateTime
116                 .Fields("sStatus").value = "pending"

118                 If lSequence = 0 Then
120                     .Fields("sequence").value = CLng(.Fields("sequence").value) + 1
                    Else
122                     .Fields("sequence").value = lSequence
                    End If
                    
124                 .Fields("sBy").value = "[" & sTitle & "] " & sDescription & " (" & sBy & ")"
126                 .Fields("sdelete").value = "false"
128                 .Fields("updates").value = supdates
130                 .Fields("noconflict").value = "false"
132                 .UpdateBatch adAffectCurrent
                    
134                 If .RecordCount > 1 Then
                    
136                     DebugPrint "(frmDynamicData.SynchHistoryEdit) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be only 1 record with sID: " & sID & " for tablename [" & sTableName & "] but there are actually " & .RecordCount & "!!!  The additional records will be deleted.  If this error message persists please contact an OASIS Developer"
                        
138                     .MoveNext

140                     Do Until .EOF
142                         .Delete adAffectCurrent
144                         .UpdateBatch adAffectCurrent
146                         .MoveNext
                        Loop
                            
                    End If
                    
148                 SynchHistoryEdit = True
            
                Else
                      
150                 DebugPrint "(frmDynamicData.SynchHistoryEdit) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be a record with sID: " & sID & " for tablename [" & sTableName & "] but it is missing!!!  This record will be created.  If this error message persists please contact an OASIS Developer"
152                 SynchHistoryAddNew sNewGUID, sID, sTitle, sDescription, sBy, sRFC3339DateTime, sTableName, bIsGeoLayer, supdates, sSynchHistPrefix
154                 SynchHistoryEdit = SynchHistoryEdit(sNewGUID, sID, sTitle, sDescription, sBy, sRFC3339DateTime, sTableName, bIsGeoLayer, supdates, sSynchHistPrefix)
        
                End If

            End If
        
        End With
        
156     Set oRS = Nothing
        SynchHistoryEdit = True
        '<EhFooter>
        Exit Function

SynchHistoryEdit_Err:
        SynchHistoryEdit = False
        Set oRS = Nothing
        MsgBox "DynamicReports.SynchHistoryEdit_Err (line " & Erl & "): " & Err.Description
        'Resume Next
        '</EhFooter>
End Function

Private Function SynchHistoryDelete(sNewGUID As String, _
                                    sID As String, _
                                    sTitle As String, _
                                    sDescription As String, _
                                    sBy As String, _
                                    sRFC3339DateTime As String, _
                                    sTableName As String, _
                                    bIsGeoLayer As Boolean, _
                                    supdates As String, _
                                    Optional sSynchHistPrefix As String = "")
        '<EhHeader>
        On Error GoTo SynchHistoryDelete_Err
        '</EhHeader>

        Dim sSynchHistoryTablename As String
        
100     SynchHistoryDelete = False
        sSynchHistoryTablename = IIf(sSynchHistPrefix = "Incidents_", "", sSynchHistPrefix) & "SynchHistory"
        
        Dim oRS As New ADODB.Recordset
102     oRS.Open "SELECT * FROM [" & sSynchHistoryTablename & "] WHERE sID = '" & sID & "' AND sTableName = '" & sTableName & "' ORDER by [sequence] DESC", mConn, adOpenDynamic, adLockBatchOptimistic

104     With oRS
 
106         If .State = adStateOpen Then
    
108             If Not .EOF Then

                    '.Fields("sID").Value = sID
110                 .Fields("sGUID").value = sNewGUID
112                 .Fields("sTableName").value = sTableName
114                 .Fields("swhen").value = sRFC3339DateTime
116                 .Fields("sStatus").value = "pending"
118                 .Fields("sequence").value = CLng(.Fields("sequence").value) + 1
120                 .Fields("sBy").value = "[" & sTitle & "] " & sDescription & " (" & sBy & ")"
122                 .Fields("sdelete").value = "true"
124                 .Fields("updates").value = supdates
126                 .Fields("noconflict").value = "false"
128                 .UpdateBatch adAffectCurrent
                    
130                 If .RecordCount > 1 Then
                    
132                      DebugPrint "(frmDynamicData.SynchHistoryDelete) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be only 1 record with sID: " & sID & " for tablename [" & sTableName & "] but there are actually " & .RecordCount & "!!!  The additional records will be deleted.  If this error message persists please contact an OASIS Developer"
                        
134                     .MoveNext

136                     Do Until .EOF
138                         .Delete adAffectCurrent
140                         .UpdateBatch adAffectCurrent
142                         .MoveNext
                        Loop
                            
                    End If
                    
144                 SynchHistoryDelete = True
            
                Else
            
146                  DebugPrint "(frmDynamicData.SynchHistoryDelete) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be a record with sID: " & sID & " for tablename [" & sTableName & "] but it is missing!!!  This record will be created.  If this error message persists please contact an OASIS Developer"
148                 SynchHistoryAddNew sNewGUID, sID, sTitle, sDescription, sBy, sRFC3339DateTime, sTableName, bIsGeoLayer, supdates, sSynchHistPrefix
150                 SynchHistoryDelete = SynchHistoryDelete(sNewGUID, sID, sTitle, sDescription, sBy, sRFC3339DateTime, sTableName, bIsGeoLayer, supdates, sSynchHistPrefix)
            
                End If
    
            End If
        
        End With
        
152     Set oRS = Nothing
        SynchHistoryDelete = True
        '<EhFooter>
        Exit Function

SynchHistoryDelete_Err:
        SynchHistoryDelete = False
        Set oRS = Nothing
        MsgBox "DynamicReports.SynchHistoryDelete_Err (line " & Erl & "): " & Err.Description
        'Resume Next
        '</EhFooter>
End Function


Private Function AsciiToBinary(Txt As String) As String

    Dim Result As String
    Dim ch As String
    Dim bin As String
    Dim i As Integer

    Txt = Replace(Txt, vbCr, "")
    Txt = Replace(Txt, vbLf, "")

    Result = ""

    For i = 1 To Len(Txt)
        ch = Mid$(Txt, i, 1)

        bin = LongToBinary(Asc(ch), False)
        Result = Result & right$(bin, 8)
    Next i

    AsciiToBinary = Result
    
End Function

Private Function BinaryToAscii(bin As String) As String

    Dim Result As String
    Dim i As Integer
    Dim next_char As String
    Dim ascii As Long

    Result = ""

    For i = 1 To Len(bin) + 18 Step 8
        next_char = Mid$(bin, i, 8)

        ascii = BinaryToLong(next_char)
        Result = Result & Chr$(ascii)
    Next i

    BinaryToAscii = Result
    
End Function

Private Function LongToBinary(ByVal long_value As Long, _
                              Optional ByVal separate_bytes As Boolean = True) As String
    Dim hex_string As String
    Dim digit_num As Integer
    Dim digit_value As Integer
    Dim nibble_string As String
    Dim result_string As String
    Dim factor As Integer
    Dim bit As Integer

    ' Convert into hex.
    hex_string = Hex$(long_value)

    ' Zero-pad to a full 8 characters.
    hex_string = right$(String$(8, "0") & hex_string, 8)

    ' Read the hexadecimal digits
    ' one at a time from right to left.
    For digit_num = 8 To 1 Step -1
        ' Convert this hexadecimal digit into a
        ' binary nibble.
        digit_value = CLng("&H" & Mid$(hex_string, digit_num, 1))

        ' Convert the value into bits.
        factor = 1
        nibble_string = ""

        For bit = 3 To 0 Step -1

            If digit_value And factor Then
                nibble_string = "1" & nibble_string
            Else
                nibble_string = "0" & nibble_string
            End If

            factor = factor * 2
        Next bit

        ' Add the nibble's string to the left of the
        ' result string.
        result_string = nibble_string & result_string
    Next digit_num

    ' Add spaces between bytes if desired.
    If separate_bytes Then
        result_string = Mid$(result_string, 1, 8) & " " & Mid$(result_string, 9, 8) & " " & Mid$(result_string, 17, 8) & " " & Mid$(result_string, 25, 8)
    End If

    ' Return the result.
    LongToBinary = result_string
End Function

Private Function BinaryToLong(ByVal binary_value As String) As Long
    Dim hex_result As String
    Dim nibble_num As Integer
    Dim nibble_value As Integer
    Dim factor As Integer
    Dim bit As Integer

    ' Remove any leading &B if present.
    ' (Note: &B is not a standard prefix, it just
    ' makes some sense.)
    binary_value = UCase$(Trim$(binary_value))

    If left$(binary_value, 2) = "&B" Then
        binary_value = Mid$(binary_value, 3)
    End If

    ' Strip out spaces in case the bytes are separated
    ' by spaces.
    binary_value = Replace(binary_value, " ", "")

    ' Left pad with zeros so we have a full 32 bits.
    binary_value = right$(String(32, "0") & binary_value, 32)

    ' Read the bits in nibbles from right to left.
    ' (A nibble is half a byte. No kidding!)
    For nibble_num = 7 To 0 Step -1
        ' Convert this nibble into a hexadecimal string.
        factor = 1
        nibble_value = 0

        ' Read the nibble's bits from right to left.
        For bit = 3 To 0 Step -1

            If Mid$(binary_value, 1 + nibble_num * 4 + bit, 1) = "1" Then
                nibble_value = nibble_value + factor
            End If

            factor = factor * 2
        Next bit

        ' Add the nibble's value to the left of the
        ' result hex string.
        hex_result = Hex$(nibble_value) & hex_result
    Next nibble_num

    ' Convert the result string into a long.
    BinaryToLong = CLng("&H" & hex_result)
End Function










