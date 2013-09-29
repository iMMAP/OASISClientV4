VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{6ADEA5C6-D7A2-11D3-AA59-00105A6F87AB}#1.0#0"; "dXDBInsp.dll"
Begin VB.UserControl DynamicDataModule 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11850
   ScaleHeight     =   8460
   ScaleWidth      =   11850
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8460
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11850
      _cx             =   20902
      _cy             =   14923
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
      Picture         =   "DynamicDataModule.ctx":0000
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   8
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
      PicturePos      =   8
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   6
      GridCols        =   5
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"DynamicDataModule.ctx":1C6D2
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CommandButton cmdSpatial 
         Caption         =   "Change Spatial Data"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5970
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   5760
      End
      Begin C1SizerLibCtl.C1Elastic lblDescription 
         Height          =   1200
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   660
         Width           =   11610
         _cx             =   20479
         _cy             =   2117
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   3
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   128
         FloodColor      =   0
         ForeColorDisabled=   -2147483631
         Caption         =   "You are currently browsing records for ........"
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
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
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   -1  'True
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
      Begin VB.CommandButton cmdEditData 
         Caption         =   "Edit Data"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8880
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.CommandButton cmdAddData 
         Caption         =   "Add Data"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5970
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   2850
      End
      Begin C1SizerLibCtl.C1Tab C1TTab1Tab2 
         Height          =   5880
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   11610
         _cx             =   20479
         _cy             =   10372
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
         BackColor       =   14737632
         ForeColor       =   0
         FrontTabColor   =   -2147483626
         BackTabColor    =   -2147483632
         TabOutlineColor =   -2147483640
         FrontTabForeColor=   32768
         Caption         =   "Browse Data|Add Data|New Tab"
         Align           =   0
         CurrTab         =   1
         FirstTab        =   0
         Style           =   5
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   3
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic667 
            Height          =   5535
            Left            =   12225
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   330
            Width           =   11580
            _cx             =   20426
            _cy             =   9763
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
            GridRows        =   9
            GridCols        =   9
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"DynamicDataModule.ctx":1C775
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.CommandButton cmdUpdateLinked 
               Caption         =   "Update"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   255
               TabIndex        =   19
               Top             =   2595
               Width           =   2085
            End
            Begin VB.CommandButton cmdAddLinked 
               Caption         =   "Add New"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   9240
               TabIndex        =   18
               Top             =   2595
               Width           =   2085
            End
            Begin C1SizerLibCtl.C1Elastic C1ElasticLinkedBottom 
               Height          =   2265
               Left            =   255
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   3000
               Width           =   11070
               _cx             =   19526
               _cy             =   3995
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
               GridRows        =   1
               GridCols        =   1
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"DynamicDataModule.ctx":1C86A
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
            End
            Begin C1SizerLibCtl.C1Elastic C1ElasticLinkedTop 
               Height          =   2265
               Left            =   255
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   270
               Width           =   11070
               _cx             =   19526
               _cy             =   3995
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
               GridRows        =   1
               GridCols        =   1
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"DynamicDataModule.ctx":1C8A2
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic666 
            Height          =   5535
            Left            =   -12195
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   330
            Width           =   11580
            _cx             =   20426
            _cy             =   9763
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
            GridCols        =   4
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"DynamicDataModule.ctx":1C8DA
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic C1ElasticBrowsing 
               Height          =   5325
               Left            =   105
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   105
               Width           =   11370
               _cx             =   20055
               _cy             =   9393
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
               GridCols        =   2
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"DynamicDataModule.ctx":1C956
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
                  Height          =   4665
                  Left            =   90
                  OleObjectBlob   =   "DynamicDataModule.ctx":1C9A2
                  TabIndex        =   17
                  Top             =   90
                  Width           =   6180
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1ElasticSimple 
            Height          =   5535
            Left            =   15
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   330
            Width           =   11580
            _cx             =   20426
            _cy             =   9763
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
            PicturePos      =   10
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   3
            GridCols        =   3
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"DynamicDataModule.ctx":1D64A
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic C1ElasticSingle 
               Height          =   5325
               Left            =   105
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   105
               Width           =   11370
               _cx             =   20055
               _cy             =   9393
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
               _GridInfo       =   $"DynamicDataModule.ctx":1D6A6
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin DXDBINSPLibCtl.dxDBInspector dxDBInspector1 
                  Height          =   5325
                  Left            =   0
                  OleObjectBlob   =   "DynamicDataModule.ctx":1D6DA
                  TabIndex        =   15
                  Top             =   0
                  Width           =   11370
               End
            End
         End
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H80000014&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         MaskColor       =   &H0080FFFF&
         TabIndex        =   3
         Top             =   7860
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H80000014&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8880
         MaskColor       =   &H0080FFFF&
         TabIndex        =   2
         Top             =   7860
         Visible         =   0   'False
         Width           =   2850
      End
      Begin CONTROLSLibCtl.dxLabel lblProgress 
         Height          =   480
         Left            =   3045
         TabIndex        =   22
         Top             =   7860
         Visible         =   0   'False
         Width           =   5775
         _Version        =   0
         _cx             =   10186
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
         Caption         =   "Dude"
         BackStyle       =   0
         BackColor       =   13160660
         ForeColor       =   0
         LabelStyle      =   3
         Label3dStyle    =   2
         Label3dOrientation=   2
         Label3dDepth    =   0
         PenWidth        =   1
         Angle           =   0
         ShadowColor     =   8421504
      End
      Begin CONTROLSLibCtl.dxProgressBar dxProgressBar1 
         Height          =   480
         Left            =   3045
         TabIndex        =   21
         Top             =   7860
         Visible         =   0   'False
         Width           =   5775
         _Version        =   65536
         _cx             =   10186
         _cy             =   847
         ForeColor       =   0
         BackColor       =   16777215
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
         ShowText        =   0   'False
         Orientation     =   0
         StartColor      =   16711680
         EndColor        =   16777215
         DrawBorderStyle =   1
         ShowTextStyle   =   0
         DrawBarStyle    =   2
         DrawBarBorderStyle=   2
      End
      Begin VB.Label lblLabel1 
         BackColor       =   &H00000080&
         Height          =   480
         Left            =   120
         TabIndex        =   20
         Top             =   7860
         Width           =   11610
      End
      Begin VB.Label lblOperation 
         BackColor       =   &H00000080&
         Caption         =   " Browsing records"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   5790
      End
   End
   Begin VB.Menu ExportOptions 
      Caption         =   "ExportOptions"
      Begin VB.Menu CopyToClipboard 
         Caption         =   "Copy to Clipboard"
      End
      Begin VB.Menu ExportToXML 
         Caption         =   "Export to XML"
      End
      Begin VB.Menu ExportToXLS 
         Caption         =   "Export to XLS"
      End
      Begin VB.Menu ExportToHTML 
         Caption         =   "Export to HTML"
      End
   End
End
Attribute VB_Name = "DynamicDataModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''
Private Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (pDst As Any, _
                                      pSrc As Any, _
                                      ByVal ByteLen As Long)
Private Declare Function SendInput _
               Lib "user32.dll" (ByVal nInputs As Long, _
                                 pInputs As Any, _
                                 ByVal cbSize As Long) As Long
Private Declare Function VkKeyScan _
               Lib "user32" _
               Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer

Private Const WM_KEYDOWN = &H100

Private Type GENERAL_INPUT
    dwType As Long
    dwData(0 To 5) As Long
End Type

Private Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    wFlags As Long
    wTime As Long
    wExtraInfo As Long
End Type

Private Type INTPOINT
    X As Integer
    Y As Integer
End Type

Private Enum EventTypes
    INPUT_MOUSE = 0
    INPUT_KEYBOARD = 1
End Enum



''''
Private Declare Sub SleepAPI Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
        
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Public Event GetSpatialLoc(oLayer As TatukGIS_XDK10.XGIS_LayerVector)
Public Event GetSpatialLocEx(oPoint As TatukGIS_XDK10.XGIS_Point)
Public Event ConvertMGRStoPT(sMGRS As String, X As Double, Y As Double)
Public Event GetCurrentExtentCentroid(oPoint As TatukGIS_XDK10.XGIS_Point)
Public Event GetNearestShapeInfo(PassedPoint As TatukGIS_XDK10.XGIS_Point, sLayerName As String, sFieldName As String, ByRef sFieldValue As String, ByRef dDistance As Double)

Private Type DYNAMIC_DATA_LINKED_RECORDSETS
    sTableName As String
    RS As ADODB.Recordset
End Type

Private Enum CurrentOperation
    browsing = 0
    Adding = 1
    Editing = 2
End Enum

Private Type DYNAMIC_DATA_DEF_SPEC_COMBORELATIONS
    sTableName As String
    sFieldName As String
    sParentFieldName As String
End Type

Private Type DYNAMIC_DATA_DEF_SPEC_VALIDATION
    sTableName As String
    sFieldName As String
    bRequired As Boolean
    sEditMask As String
    sValidation As String
End Type

Private Type DYNAMIC_DATA_DEF_SPEC_NEARBYFEATURES
    sLayerName As String
    sLayerFieldName As String
    sDataEntryTableName As String
    sDataEntryFieldName As String
    bDistance As Boolean
End Type

Private Type DYNAMIC_DATA_DEF_SPEC_AUTOCALCFIELDS
    sTableName As String
    sFieldNameChild As String
    sFieldNameParent As String
End Type

Private Type DYNAMIC_DATA_DEF_SPEC_DETAILED
    lRank As Long
    sTableName As String
    sCaption As String
    sDescription As String
    lDescFontSize As Long
    sDataEntryFields As String
    sLockedFields As String
    bIsMasterTable As Boolean
    bIsLinkedTable As Boolean
    sSQLView As String
    sSQLViewMSSQL As String
End Type

Private Type DYNAMIC_DATA_DEF_SPEC
    lMaxRank As Long
    lCurrentRank As Long
    Detail() As DYNAMIC_DATA_DEF_SPEC_DETAILED
End Type

Private Type DYNAMIC_DATA_DEF_TABLES
    TableName As String
    AllowRead As Boolean
    AllowAppend As Boolean
    AllowEdit As Boolean
    AllowDelete As Boolean
    IsMasterTable As Boolean
    IsLinkedTable As Boolean
    IsDDTable As Boolean
    IsGEOTable As Boolean
    ListIndex As Long
End Type

Private Type DYNAMIC_DATA_DEF
    Name As String
    desc As String
    Tables() As DYNAMIC_DATA_DEF_TABLES
    NumOfLinkedTables As Long
    NumOfDDTables As Long
    NumMainDataEntryPages As Long
    LinkedTableNames As String
    DDTableNames As String
    ListIndex As Long
    Prefix As String
    ConnectionString As String
    ExcludedFields  As String
    LockedFields As String
    Specification As DYNAMIC_DATA_DEF_SPEC
End Type

Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Private mDDRS As ADODB.Recordset
Private mDDRSManual As ADODB.Recordset
Private mDDRSLinked As ADODB.Recordset
Private mDDRSLinkedManual As ADODB.Recordset

Private mDDRSGuid As String
Private mDDRSLinkedGuid As String

Private sFilterBeforeAdd As String

Private mConn As ADODB.Connection
Private mLayer As TatukGIS_XDK10.XGIS_LayerSqlAdo
Private mLayerCopy As TatukGIS_XDK10.XGIS_LayerVector
Private mShape As TatukGIS_XDK10.XGIS_Shape
Private mShapeNew As TatukGIS_XDK10.XGIS_Shape

Private bLinkedTableActive As Boolean
Private bLoadOfmDDRSInProgess As Boolean
Private bDisableEvents As Boolean
Private bDisableEventsDuringScroll As Boolean

Private DDDefs() As DYNAMIC_DATA_DEF
Private DDValidation() As DYNAMIC_DATA_DEF_SPEC_VALIDATION
Private DDComboRelations() As DYNAMIC_DATA_DEF_SPEC_COMBORELATIONS
Private DDAutoCalcFields() As DYNAMIC_DATA_DEF_SPEC_AUTOCALCFIELDS
Private DDNearbyFeatures() As DYNAMIC_DATA_DEF_SPEC_NEARBYFEATURES
Private mCreatedPoint As XGIS_Point

Private DDDefCurrent As DYNAMIC_DATA_DEF
Private DDTableCurrent As DYNAMIC_DATA_DEF_TABLES
Private DDTableCurrentDetail As DYNAMIC_DATA_DEF_SPEC_DETAILED
Private sLinkedTableNames() As String
Private sKMLString As String
Private AllDDDefTables() As DYNAMIC_DATA_DEF_TABLES
Private mDDRSLinkedCollection() As DYNAMIC_DATA_LINKED_RECORDSETS

Public listDatabases As ListBox
Public ListDataElements As ListBox
Private bCommitShape As Boolean
Private eCurrentOperation As CurrentOperation

Private SavedId As String
Private PrevEnumVal As String
Private SavedIdAssigned As Boolean
Private lMaxProgressBarSteps As Long
Private lProgressBarStep As Long

''''
Private Sub SimulateRightArrow()

    Dim GenInput As GENERAL_INPUT
    Dim tmp As INTPOINT

    GenInput.dwType = EventTypes.INPUT_KEYBOARD

    tmp.X = vbKeyRight
    tmp.Y = 0
    CopyMemory GenInput.dwData(0), tmp, Len(tmp)
    GenInput.dwData(1) = WM_KEYDOWN
    GenInput.dwData(2) = 0
    GenInput.dwData(3) = 0

    SendInput 1, GenInput, Len(GenInput)

End Sub

Public Sub CommitShape(bDoCommitShape As Boolean)
    bCommitShape = bDoCommitShape
End Sub

Public Function IsEditing() As Boolean

    If Not eCurrentOperation = browsing And dxDBInspector1.Visible Then
        IsEditing = True
    Else
        IsEditing = False
    End If

End Function

Private Sub GetNearbyFeatures(oCn As ADODB.Connection, _
                              sPreFix As String)
        '<EhHeader>
        On Error GoTo GetNearbyFeatures_Err
        '</EhHeader>
    
        Dim oRS As New ADODB.Recordset
        Dim i As Long
         
100     ReDim DDNearbyFeatures(0)
    
102     With oRS
104         .Open "SELECT * FROM [" & sPreFix & "NearbyFeatures]", oCn, adOpenDynamic, adLockBatchOptimistic

106         If Not .State = 0 Then
108             ReDim DDNearbyFeatures(0)
            
110             Do Until .EOF
112                 ReDim Preserve DDNearbyFeatures(UBound(DDNearbyFeatures) + 1)
114                 DDNearbyFeatures(i).sDataEntryFieldName = .Fields("sDataEntryFieldName").value
116                 DDNearbyFeatures(i).sLayerName = .Fields("sLayerName").value
                    DDNearbyFeatures(i).sLayerFieldName = .Fields("sLayerFieldName").value
118                 DDNearbyFeatures(i).sDataEntryTableName = .Fields("sDataEntryTableName").value
                    DDNearbyFeatures(i).bDistance = .Fields("bCalculateDistance").value
120                 i = i + 1
122                 .MoveNext
                Loop

124             .Close
            End If

        End With

126     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

GetNearbyFeatures_Err:
        Set oRS = Nothing
        '</EhFooter>
End Sub

Private Sub GetComboRelations(oCn As ADODB.Connection, _
                              sPreFix As String)
        '<EhHeader>
        On Error GoTo GetComboRelations_Err
        '</EhHeader>
    
        Dim oRS As New ADODB.Recordset
        Dim i As Long
         
100     ReDim DDComboRelations(0)
    
102     With oRS
104         .Open "SELECT * FROM [" & sPreFix & "ComboRelations]", oCn, adOpenDynamic, adLockBatchOptimistic

106         If Not .State = 0 Then
108             ReDim DDComboRelations(0)
            
110             Do Until .EOF
112                 ReDim Preserve DDComboRelations(UBound(DDComboRelations) + 1)
114                 DDComboRelations(i).sFieldName = .Fields("sFieldName").value
118                 DDComboRelations(i).sParentFieldName = .Fields("sParentFieldName").value
120                 DDComboRelations(i).sTableName = .Fields("sTableName").value
122                 i = i + 1
124                 .MoveNext
                Loop

126             .Close
            End If

        End With

128     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

GetComboRelations_Err:
        Set oRS = Nothing
        '</EhFooter>
End Sub

Private Sub GetValidation(oCn As ADODB.Connection, _
                          sPreFix As String)
        '<EhHeader>
        On Error GoTo GetValidation_Err
        '</EhHeader>
    
        Dim oRS As New ADODB.Recordset
        Dim i As Long
         
100     ReDim DDValidation(0)
    
102     With oRS
104         .Open "SELECT * FROM [" & sPreFix & "Validation]", oCn, adOpenDynamic, adLockBatchOptimistic

106         If Not .State = 0 Then
108             ReDim DDValidation(0)
            
110             Do Until .EOF
112                 ReDim Preserve DDValidation(UBound(DDValidation) + 1)
114                 DDValidation(i).sTableName = .Fields("sDataEntryTableName").value
116                 DDValidation(i).sFieldName = .Fields("sDataEntryFieldName").value

118                 If Not IsNull(.Fields("bRequired").value) Then DDValidation(i).bRequired = .Fields("bRequired").value
120                 If Not IsNull(.Fields("sEditMask").value) Then DDValidation(i).sEditMask = .Fields("sEditMask").value
122                 If Not IsNull(.Fields("sValidation").value) Then DDValidation(i).sValidation = .Fields("sValidation").value
124                 i = i + 1
126                 .MoveNext
                Loop

128             .Close
            End If

        End With

130     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

GetValidation_Err:
        Set oRS = Nothing
        '</EhFooter>
End Sub

Private Sub GetAutoCalcFields(oCn As ADODB.Connection, _
                              sPreFix As String)
        '<EhHeader>
        On Error GoTo GetAutoCalcFields_Err
        '</EhHeader>
    
        Dim oRS As New ADODB.Recordset
        Dim i As Long
        
100     ReDim DDAutoCalcFields(0)
    
102     With oRS
104         .Open "SELECT * FROM [" & sPreFix & "AutoCalcFields]", oCn, adOpenDynamic, adLockBatchOptimistic

106         If Not .State = 0 Then
108             ReDim DDAutoCalcFields(0)
            
110             Do Until .EOF
112                 ReDim Preserve DDAutoCalcFields(UBound(DDAutoCalcFields) + 1)
114                 DDAutoCalcFields(i).sTableName = .Fields("sTableName").value
118                 DDAutoCalcFields(i).sFieldNameChild = .Fields("sFieldNameChild").value
                    DDAutoCalcFields(i).sFieldNameParent = .Fields("sFieldNameParent").value
122                 i = i + 1
124                 .MoveNext
                Loop

126             .Close
            End If

        End With

128     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

GetAutoCalcFields_Err:
        Set oRS = Nothing
        '</EhFooter>
End Sub

Private Function GetComboRelationChildField(sTableName As String, _
                                            sFieldName As String) As String

    Dim i As Long
    i = UBound(DDComboRelations)
    GetComboRelationChildField = ""
    
    Do Until i < 0 Or Not GetComboRelationChildField = ""

        If DDComboRelations(i).sTableName = sTableName And DDComboRelations(i).sParentFieldName = sFieldName Then GetComboRelationChildField = DDComboRelations(i).sFieldName
        i = i - 1
    Loop

End Function

Private Function SynchHistoryAddNew(sNewGUID As String, _
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
102     oRS.Open "SELECT * FROM [" & sSynchHistPrefix & "SynchHistory] WHERE sID = '" & sID & "' AND sTableName = '" & sTableName & "'", mConn, adOpenDynamic, adLockBatchOptimistic

104     With oRS
 
106         If .State = adStateOpen Then
    
108             If .RecordCount > 0 Then
                    
110                 DebugPrint "(DynamicDataModule.SynchHistoryAddNew) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be 0 record with sID: " & sID & " for tablename [" & sTableName & "] but there are actually " & .RecordCount & "!!!  These records will be deleted.  If this error message persists please contact an OASIS Developer"

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
142                 .Fields("noconflict").value = "local"
144                 .UpdateBatch adAffectCurrent
146                 SynchHistoryAddNew = True
            
                End If
    
            Else
            
148             DebugPrint "(DynamicDataModule.SynchHistoryAddNew) Table [" & sSynchHistPrefix & "SynchHistory] failed to open"
            
            End If
        
        End With
        
150     Set oRS = Nothing
        SynchHistoryAddNew = True
        '<EhFooter>
        Exit Function

SynchHistoryAddNew_Err:
        SynchHistoryAddNew = False
        Set oRS = Nothing
        MsgBox "DynamicDataModule.SynchHistoryAddNew_Err (line " & Erl & "): " & Err.Description
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
                                  Optional lSequence As Long = 0) As Boolean
        '<EhHeader>
        On Error GoTo SynchHistoryEdit_Err
        '</EhHeader>

100     SynchHistoryEdit = False
    
        Dim oRS As New ADODB.Recordset
102     oRS.Open "SELECT * FROM [" & sSynchHistPrefix & "SynchHistory] WHERE sID = '" & sID & "' AND sTableName = '" & sTableName & "' ORDER by [sequence] DESC", mConn, adOpenDynamic, adLockBatchOptimistic

104     With oRS
 
106         If .State = adStateOpen Then
    
108             If Not .EOF Then

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
130                 '.Fields("noconflict").value = "edit"
132                 .UpdateBatch adAffectCurrent
                    
134                 If .RecordCount > 1 Then
                    
136                     DebugPrint "(DynamicDataModule.SynchHistoryEdit) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be only 1 record with sID: " & sID & " for tablename [" & sTableName & "] but there are actually " & .RecordCount & "!!!  The additional records will be deleted.  If this error message persists please contact an OASIS Developer"
                        
138                     .MoveNext

140                     Do Until .EOF
142                         .Delete adAffectCurrent
144                         .UpdateBatch adAffectCurrent
146                         .MoveNext
                        Loop
                            
                    End If
                    
148                 SynchHistoryEdit = True
            
                Else
                      
150                 DebugPrint "(DynamicDataModule.SynchHistoryEdit) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be a record with sID: " & sID & " for tablename [" & sTableName & "] but it is missing!!!  This record will be created.  If this error message persists please contact an OASIS Developer"
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
        MsgBox "DynamicDataModule.SynchHistoryEdit_Err (line " & Erl & "): " & Err.Description
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
                                    Optional sSynchHistPrefix As String = "") As Boolean
        '<EhHeader>
        On Error GoTo SynchHistoryDelete_Err
        '</EhHeader>

100     SynchHistoryDelete = False
        
        Dim oRS As New ADODB.Recordset
102     oRS.Open "SELECT * FROM [" & sSynchHistPrefix & "SynchHistory] WHERE sID = '" & sID & "' AND sTableName = '" & sTableName & "' ORDER by [sequence] DESC", mConn, adOpenDynamic, adLockBatchOptimistic

104     With oRS
 
106         If .State = adStateOpen Then
    
108             If Not .EOF Then

110                 .Fields("sGUID").value = sNewGUID
112                 .Fields("sTableName").value = sTableName
114                 .Fields("swhen").value = sRFC3339DateTime
116                 .Fields("sStatus").value = "pending"
118                 .Fields("sequence").value = CLng(.Fields("sequence").value) + 1
120                 .Fields("sBy").value = "[" & sTitle & "] " & sDescription & " (" & sBy & ")"
122                 .Fields("sdelete").value = "true"
124                 .Fields("updates").value = supdates
126                 '.Fields("noconflict").value = "deletion"
128                 .UpdateBatch adAffectCurrent
                    
130                 If .RecordCount > 1 Then
                    
132                     DebugPrint "(DynamicDataModule.SynchHistoryDelete) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be only 1 record with sID: " & sID & " for tablename [" & sTableName & "] but there are actually " & .RecordCount & "!!!  The additional records will be deleted.  If this error message persists please contact an OASIS Developer"
                        
134                     .MoveNext

136                     Do Until .EOF
138                         .Delete adAffectCurrent
140                         .UpdateBatch adAffectCurrent
142                         .MoveNext
                        Loop
                            
                    End If
                    
144                 SynchHistoryDelete = True
            
                Else
            
146                 DebugPrint "(DynamicDataModule.SynchHistoryDelete) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be a record with sID: " & sID & " for tablename [" & sTableName & "] but it is missing!!!  This record will be created.  If this error message persists please contact an OASIS Developer"
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
        MsgBox "DynamicDataModule.SynchHistoryDelete_Err (line " & Erl & "): " & Err.Description
        'Resume Next
        '</EhFooter>
End Function

Private Sub DisplayGridData()
        '<EhHeader>
        On Error GoTo DisplayGridData_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
        Dim oRSView As New ADODB.Recordset
        Dim i As Long
        
100     Set oRS = IIf(DDTableCurrent.IsLinkedTable, mDDRSLinked, mDDRS)

102     lblDescription.caption = "please wait momentarily for the data to load............"
104     lblDescription.FontSize = 15
106     lblDescription.Refresh
    
108     ScanForDDFields
110     dxDBInspector1.DividerPos = dxDBInspector1.Width / (1.75 * Screen.TwipsPerPixelX)
        
112     If (C1TTab1Tab2.CurrTab = 0 Or DDTableCurrent.IsLinkedTable) Then

114         If Not oRS.EOF Or Not oRS.Bof Then oRS.MoveFirst
116         If DDTableCurrent.IsLinkedTable Then
118             dxDBGrid1.KeyField = "GUID2"
            Else
120             dxDBGrid1.KeyField = "GUID1"
            End If
            
122         i = 0

124         Do Until i >= UBound(DDDefCurrent.Specification.Detail)
            
126             If DDDefCurrent.Specification.Detail(i).sTableName = DDTableCurrent.TableName Then
                    
128                 dxDBGrid1.Columns.DestroyColumns

130                 If bSQLServerInUse Then
132                     oRSView.Open DDDefCurrent.Specification.Detail(i).sSQLViewMSSQL, mConn, adOpenDynamic, adLockBatchOptimistic
                    Else
134                     oRSView.Open DDDefCurrent.Specification.Detail(i).sSQLView, mConn, adOpenDynamic, adLockBatchOptimistic
                    End If
                    
136                 If Not DDTableCurrent.IsLinkedTable Then
                    
138                     Set dxDBGrid1.DataSource = oRSView
                        
                    Else
                        
140                     Set dxDBGrid1.DataSource = oRS
                        
                    End If
                    
142                 dxDBGrid1.Columns.RetrieveFields
144                 i = UBound(DDDefCurrent.Specification.Detail)

                Else
146                 i = i + 1
                End If
                
            Loop
            
148         i = 0

150         Do Until i >= dxDBGrid1.Columns.Count
            
152             If dxDBGrid1.Columns(i).FieldName = "UID" Then dxDBGrid1.Columns(i).Visible = False
154             If dxDBGrid1.Columns(i).FieldName = "GUID1" Then dxDBGrid1.Columns(i).Visible = False
156             If dxDBGrid1.Columns(i).FieldName = "GUID2" Then dxDBGrid1.Columns(i).Visible = False

158             If DDTableCurrent.IsLinkedTable And InStr(oRSView.Source, dxDBGrid1.Columns(i).FieldName) < 1 Then
160                 dxDBGrid1.Columns(i).Visible = False
                End If
            
162             If DDTableCurrent.IsLinkedTable And left(dxDBGrid1.Columns(i).FieldName, 3) = "dd_" Then
164                 SetGridLookup dxDBGrid1.Columns(i).FieldName
                End If

166             i = i + 1
            Loop
               
168         If DDTableCurrent.IsLinkedTable Then
                    
170             Set dxDBGrid1.DataSource = Nothing
172             Set dxDBGrid1.DataSource = oRS
                
                Dim jj
                For jj = 0 To dxDBGrid1.Columns.Count - 1
                    If dxDBGrid1.Columns(jj).Visible Then
                        dxDBGrid1.Columns(jj).Sorted = csUp
                        Exit For
                    End If
                Next
                    
            End If
        
        End If

174     Call SetGridCaptions
176     Call SetAccessRights
        
        '<EhFooter>
        Exit Sub

DisplayGridData_Err:
        MsgBox "DynamicDataModule.DisplayGridData_Err (line " & Erl & "): " & Err.Description
        Resume Next
        '</EhFooter>
End Sub

Public Sub ListDatabases_Click()
        '<EhHeader>
        On Error GoTo ListDatabases_Click_Err
        '</EhHeader>

        Dim i As Long
        Dim lCountOfTables As Long
        Dim lTableIndex As Long
        Dim sTableName As String
        Dim oRSCombo As ADODB.Recordset
                
100     C1TTab1Tab2.Visible = False
102     cmdAddData.Visible = False
104     cmdEditData.Visible = False
106     lblOperation.caption = "  " & listDatabases.Text
108     lblDescription.caption = "Please select a module element on the left..."
110     lblDescription.FontSize = 15
        
112     Call PointAtCurrentDDInList
114     ListDataElements.Clear
    
116     Set mConn = New ADODB.Connection

118     If bSQLServerInUse Then
120         mConn.Open g_sGlobalConnectionString
            GetDDDetailedSpec DDDefCurrent.Prefix, DDDefCurrent, g_sGlobalConnectionString
        Else

122         mConn.ConnectionString = DDDefCurrent.ConnectionString
124         mConn.ConnectionString = Replace(mConn.ConnectionString, "\data\db\dynamicdata", AppPath & "data\db\dynamicdata", , , vbTextCompare)
126         mConn.CursorLocation = g_sGlobalCursorLocation

            GetDDDetailedSpec DDDefCurrent.Prefix, DDDefCurrent, mConn.ConnectionString

            On Error GoTo NoSuchDatabase
128         mConn.Open
            On Error GoTo ListDatabases_Click_Err
        
        End If
    
130     lCountOfTables = UBound(DDDefCurrent.Tables)
132     i = 1
    
134     lTableIndex = 0

136     Do Until i = lCountOfTables + 1
    
138         With DDDefCurrent.Tables(i)
        
140             .ListIndex = 666666
    
142             If Not right$(.TableName, 4) = "_GEO" And .AllowRead And Not left$(.TableName, 4) = "link" Then
        
144                 If .IsLinkedTable Then
146                     sTableName = .TableName
148                 ElseIf .IsMasterTable Then
150                     sTableName = "DATA ENTRY"
                    Else
152                     sTableName = right$(.TableName, Len(.TableName) - 2)
                    End If

154                 If .IsGEOTable Then sTableName = left$(sTableName, Len(sTableName) - 4)
                
156                 ListDataElements.AddItem IIf(Len(GetCaption(.TableName)) > 1, GetCaption(.TableName), sTableName), lTableIndex
158                 .ListIndex = lTableIndex
160                 lTableIndex = lTableIndex + 1
        
                End If

            End With

162         i = i + 1
        Loop
        
164     GetComboRelations mConn, DDDefCurrent.Prefix
166     GetAutoCalcFields mConn, DDDefCurrent.Prefix
168     GetNearbyFeatures mConn, DDDefCurrent.Prefix
170     GetValidation mConn, DDDefCurrent.Prefix
        
172     'Set C1Elastic1.Picture = dxDBInspector1.BackgroundBitmap
174     C1Elastic1.BackColor = vbWhite
176     'Set lblDescription.Picture = Nothing
        
        '        ReDim DDComboRelations(0)
        '
        '        If DoesTableExist(mConn.ConnectionString, DDDefCurrent.Prefix & "ComboRelations") Then
        '
        '            Set oRSCombo = New adodb.Recordset
        '            oRSCombo.Open "SELECT * FROM [" & DDDefCurrent.Prefix & "ComboRelations]", mConn, adOpenDynamic, adLockBatchOptimistic
        '
        '            Do Until oRSCombo.EOF
        '                ReDim Preserve DDComboRelations(UBound(DDComboRelations) + 1)
        '                DDComboRelations(UBound(DDComboRelations)).sFieldName = oRSCombo.Fields("sFieldName").Value
        '                DDComboRelations(UBound(DDComboRelations)).sParentFieldName = oRSCombo.Fields("sParentFieldName").Value
        '                DDComboRelations(UBound(DDComboRelations)).sComboQuery = oRSCombo.Fields("sComboQuery").Value
        '                DDComboRelations(UBound(DDComboRelations)).bDisableBeforeParent = oRSCombo.Fields("bDisableBeforeParent").Value
        '                oRSCombo.MoveNext
        '            Loop
        '
        '            oRSCombo.Close
        '            Set oRSCombo = Nothing
        '
        '        End If
178     lblLabel1.Visible = True
180     dxProgressBar1.Visible = False
182     lblProgress.Visible = False
        Exit Sub

NoSuchDatabase:
184     MsgBox "This database is not available.", vbInformation
        'Resume Next
        
        '<EhFooter>
        Exit Sub

ListDatabases_Click_Err:
        'MsgBox "DynamicDataModule.ListDatabases_Click_Err (" & Erl & " ) " & Err.Description
        MsgBox "The module you selected is in error.  This might be due to the database schema being synchronised.  Please leave OASIS open for a while to fully synchronise and then restart OASIS.  If this does not solve the issue then please contact an OASIS administrator.", vbInformation, "Config error"
        
        '</EhFooter>
End Sub

Private Sub UpdateProgressBarForNewTable()

    If DDTableCurrent.IsDDTable Then
        dxProgressBar1.Pos = 0
        dxProgressBar1.Step = 100 / 2
        dxProgressBar1.DoStep
        lMaxProgressBarSteps = 1
        lProgressBarStep = 1
    Else
        dxProgressBar1.Pos = 0
        dxProgressBar1.Step = 100 / (DDDefCurrent.NumMainDataEntryPages + 1)
        dxProgressBar1.DoStep
        lMaxProgressBarSteps = DDDefCurrent.NumMainDataEntryPages
        lProgressBarStep = 1
    End If
    
    lblProgress.caption = "Step " & lProgressBarStep & " of " & lMaxProgressBarSteps
    dxProgressBar1.Visible = True
    lblProgress.Visible = True
        
End Sub

Private Sub cmdAddData_Click()
        '<EhHeader>
        On Error GoTo cmdAddData_Click_Err
        '</EhHeader>
        dxDBInspector1.BackColor = vbWhite
100     C1TTab1Tab2.CurrTab = 1
102     C1TTab1Tab2.Visible = False
104     cmdEditData.Enabled = False
106     cmdAddData.Enabled = False
108     ListDataElements.Enabled = False
110     listDatabases.Enabled = False
112     cmdCancel.Visible = True
114     lblOperation.caption = "  Adding data..."
118     eCurrentOperation = Adding

120     mDDRSGuid = GUIDGen
122     mDDRS.AddNew Array("GUID1"), Array(mDDRSGuid)
124     mDDRS.Fields("GUID1").value = mDDRSGuid
126     CopyRSValues mDDRS, mDDRSManual
        UpdateLinkedTableCollection

128     If DDTableCurrent.IsGEOTable Then
        
            Dim oPoint As New TatukGIS_XDK10.XGIS_Point
130         RaiseEvent GetCurrentExtentCentroid(oPoint)
132         Set mShape = mLayer.CreateShape(XgisShapeTypeUnknown)
134         mShape.AddPart
136         mShape.AddPoint oPoint
        
138         cmdSpatial.Enabled = True
140         cmdSpatial.Visible = True
142         mDDRS.Fields("UID").value = mShape.uID
        
        End If

146     Call UpdateCaptionOfNextButton
148     cmdNext.Visible = True
150     C1TTab1Tab2.Visible = True
        dxDBInspector1.Visible = True
        'dxProgressBar1.StartColor=ubound
        'MsgBox UBound(mDDRSLinkedCollection)
        UpdateValidationMask DDTableCurrent.TableName
        UpdateProgressBarForNewTable
        lblLabel1.Visible = False
            
        '<EhFooter>
        Exit Sub

cmdAddData_Click_Err:
        MsgBox "DynamicDataModule.cmdAddData_Click_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Sub cmdAddLinked_Click()
        '<EhHeader>
        On Error GoTo cmdAddLinked_Click_Err
        '</EhHeader>
        Dim i As Long
        dxDBInspector1.PostEditor
100     mDDRSLinked.AddNew Array("GUID1", "GUID2"), Array(mDDRSGuid, mDDRSLinkedGuid)
              
102     CopyRSValues mDDRSLinkedManual, mDDRSLinked, True

        If DDTableCurrent.TableName = "linkItemsDisAndRec" Then NomadTotalPriceCalc mDDRSLinked
104     mDDRSLinkedGuid = GUIDGen
        mDDRSLinked.Fields("GUID2").value = mDDRSLinkedGuid
        
        If SafeMoveFirst(mDDRSLinked) And mDDRSLinked.RecordCount > 0 Then
            CopyRSValues mDDRSLinked, mDDRSLinkedManual, True
            mDDRSLinkedGuid = mDDRSLinked.Fields("GUID2").value
        End If
        
        Exit Sub

106     If FieldExists(mDDRSLinkedManual, "GUID1") And Not FieldExists(mDDRSLinkedManual, "GUID2") Then
108         mDDRSLinkedManual.AddNew Array("GUID1"), Array(mDDRSGuid)
110     ElseIf FieldExists(mDDRSLinkedManual, "GUID2") And Not FieldExists(mDDRSLinkedManual, "GUID1") Then
112         mDDRSLinkedManual.AddNew Array("GUID2"), Array(mDDRSLinkedGuid)
114     ElseIf FieldExists(mDDRSLinkedManual, "GUID1") And FieldExists(mDDRSLinkedManual, "GUID2") Then
116         mDDRSLinkedManual.AddNew Array("GUID1", "GUID2"), Array(mDDRSGuid, mDDRSLinkedGuid)
        Else
118         mDDRSLinkedManual.AddNew
        End If
        
        '<EhFooter>
        Exit Sub

cmdAddLinked_Click_Err:
        MsgBox "DynamicDataModule.cmdAddLinked_Click_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Sub NomadTotalPriceCalc(ByRef oRSPassed As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo NomadTotalPriceCalc_Err
        '</EhHeader>

        Dim oRSLookup As New ADODB.Recordset
100     oRSLookup.Open "SELECT [item_unitprice] FROM [dd_NOMADCORE_ddItems] WHERE [GUID1] = '" & oRSPassed.Fields("dd_NOMADCORE_ddItems").value & "'", mConn, adOpenDynamic, adLockBatchOptimistic
    
102     If oRSLookup.State = adStateOpen Then
    
104         If Not oRSLookup.EOF Then
    
106             If Len(CStr(oRSLookup.Fields(0).value)) > 0 Then
            
108                 If oRSLookup.Fields(0).value > 0 Then
110                     oRSPassed.Fields("item_totalprice").value = oRSLookup.Fields(0).value * oRSPassed.Fields("item_requestqty").value
                    End If
                
                End If
    
            End If
    
        End If
    
        '<EhFooter>
        Exit Sub

NomadTotalPriceCalc_Err:
        DebugPrint "DynamicDataModule.NomadTotalPriceCalc_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
        
        Dim lGUID As String
        
100     dxDBInspector1.Visible = False
102     cmdNext.Visible = False
104     cmdCancel.Visible = False
106     lblDescription.caption = "please wait momentarily for the form to load..."
108     lblDescription.FontSize = 15

110     DoEvents
        
112     If cmdCancel.caption = "Back" Then
        
114         C1TTab1Tab2.Visible = False
        
116         If DDDefCurrent.Specification.Detail(DDDefCurrent.Specification.lCurrentRank - 1).bIsMasterTable Then
118             CopyRSValues mDDRSManual, mDDRS
            End If
        
120         DDDefCurrent.Specification.lCurrentRank = DDDefCurrent.Specification.lCurrentRank - 1
122         MoveToNextTable DDDefCurrent.Specification.Detail(DDDefCurrent.Specification.lCurrentRank - 1).sTableName
124         UpdateContainerOfControls True
            
126         If DDDefCurrent.Specification.lCurrentRank = 1 Then
128             cmdCancel.caption = "Cancel"
            End If
            
130         cmdNext.caption = "Next"
132         C1TTab1Tab2.Visible = True
134         cmdNext.Visible = True
136         cmdCancel.Visible = True
            dxDBInspector1.Visible = True
            UpdateDescription
            UpdateValidationMask DDTableCurrent.TableName
            dxProgressBar1.DoStepBy -dxProgressBar1.Step
            lProgressBarStep = lProgressBarStep - 1
            lblProgress.caption = "Step " & lProgressBarStep & " of " & lMaxProgressBarSteps
        Else
        
138         If DDTableCurrent.IsGEOTable Then
        
140             If eCurrentOperation = Editing Then
142                 mShape.Reset
144                 mLayer.RevertAll
146             ElseIf eCurrentOperation = Adding Then
148                 mLayer.Delete mDDRS.Fields("UID").value
150                 mShape.Reset
152                 mLayer.SaveAll
154                 mLayer.RevertAll
                End If
            
156             Set mShapeNew = Nothing
158             Set mLayerCopy = Nothing
160             sKMLString = ""
        
            End If
        
164         cmdCancel.Visible = False
            lGUID = mDDRS.Fields("GUID1").value
166         Call ListDataElements_Click

            If Not eCurrentOperation = Adding Then
                If Not mDDRS.EOF Or Not mDDRS.Bof Then
                    mDDRS.MoveFirst
                    mDDRS.Find "[GUID1] = '" & lGUID & "'"
                    dxDBGrid1.Dataset.Locate "GUID1", lGUID, False, False
                End If
            End If
            
            eCurrentOperation = browsing
168         UpdateDescription
            dxProgressBar1.Visible = False
            lblProgress.Visible = False
            lblLabel1.Visible = False
        End If
        
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox "DynamicDataModule.cmdCancel_Click_Err (" & Erl & " ) " & Err.Description
        dxDBInspector1.Visible = True
        cmdNext.Visible = True
        cmdCancel.Visible = True
        '</EhFooter>
End Sub

Private Sub MoveToNextTable(sTableName As String)
        '<EhHeader>
        On Error GoTo MoveToNextTable_Err
        '</EhHeader>

        Dim sSQL As String
        Dim oRS As ADODB.Recordset
    
100     bDisableEvents = True
102     PointAtCurrentTable sTableName
104     dxDBInspector1.ClearRows
    
106     If DDTableCurrent.IsLinkedTable Then
    
108         If GetRSFromLinkedTablesCollection(DDTableCurrent.TableName) Is Nothing Then
        
110             Set mDDRSLinked = New ADODB.Recordset
112             Set mDDRSLinkedManual = New ADODB.Recordset
114             mDDRSLinked.Open "SELECT * FROM [" & DDDefCurrent.Prefix & DDTableCurrent.TableName & "] WHERE [GUID1] = '" & mDDRSGuid & "'", mConn, adOpenDynamic, adLockBatchOptimistic
116             SetRSInLinkedTablesCollection DDTableCurrent.TableName, mDDRSLinked
        
118         ElseIf GetRSFromLinkedTablesCollection(DDTableCurrent.TableName).State = adStateClosed Then
       
120             Set mDDRSLinked = New ADODB.Recordset
122             Set mDDRSLinkedManual = New ADODB.Recordset
124             mDDRSLinked.Open "SELECT * FROM [" & DDDefCurrent.Prefix & DDTableCurrent.TableName & "] WHERE [GUID1] = '" & mDDRSGuid & "'", mConn, adOpenDynamic, adLockBatchOptimistic
126             SetRSInLinkedTablesCollection DDTableCurrent.TableName, mDDRSLinked
       
            Else
       
128             Set oRS = GetRSFromLinkedTablesCollection(DDTableCurrent.TableName)
130             Set mDDRSLinked = oRS.Clone
        
            End If
        
132         Set mDDRSLinkedManual = CreateCustomRS(mDDRSLinked)
134         Set dxDBInspector1.DataSource = mDDRSLinkedManual
        
        Else
    
136         Set mDDRSManual = CreateCustomRS(mDDRS)
138         Set dxDBInspector1.DataSource = mDDRSManual
        
        End If
    
140     DisplayGridData
142     C1Elastic1.Refresh
144     lblOperation.caption = "  Browsing records..."
146     Call SetAccessRights
148     bDisableEvents = False
    
        '<EhFooter>
        Exit Sub

MoveToNextTable_Err:
        MsgBox "DynamicDataModule.MoveToNextTable_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Sub PointAtCurrentTable(sTableName As String)
        '<EhHeader>
        On Error GoTo PointAtCurrentTable_Err
        '</EhHeader>

        Dim i As Long
100     i = 0
    
102     Do Until i = UBound(DDDefCurrent.Tables) + 1
    
104         If DDDefCurrent.Tables(i).TableName = sTableName Then
106             DDTableCurrent = DDDefCurrent.Tables(i)
            End If

108         i = i + 1
        Loop
     
        i = 0
    
        Do Until i = UBound(DDDefCurrent.Specification.Detail) + 1
    
            If DDDefCurrent.Specification.Detail(i).sTableName = sTableName Then
                DDTableCurrentDetail = DDDefCurrent.Specification.Detail(i)
            End If

            i = i + 1
        Loop

        '<EhFooter>
        Exit Sub

PointAtCurrentTable_Err:
        MsgBox "DynamicDataModule.PointAtCurrentTable_Err (" & Erl & " ) " & Err.Description
        
        '</EhFooter>
End Sub

Private Sub PointAtCurrentTableInList()
        '<EhHeader>
        On Error GoTo PointAtCurrentTableInList_Err
        '</EhHeader>

        Dim i As Long
100     i = 0
    
102     Do Until i = UBound(DDDefCurrent.Tables) + 1
    
104         If DDDefCurrent.Tables(i).ListIndex = ListDataElements.ListIndex Then
106             DDTableCurrent = DDDefCurrent.Tables(i)

                If DDTableCurrent.IsMasterTable Then DDDefCurrent.Specification.lCurrentRank = 1
            End If

108         i = i + 1
        Loop

        '<EhFooter>
        Exit Sub

PointAtCurrentTableInList_Err:
        MsgBox "DynamicDataModule.PointAtCurrentTableInList_Err (" & Erl & " ) " & Err.Description
        
        '</EhFooter>
End Sub

Private Function GetCaption(sTableName As String)

    Dim i As Long
    i = 0
    
    Do Until i = UBound(DDDefCurrent.Specification.Detail) + 1

        If DDDefCurrent.Specification.Detail(i).sTableName = sTableName Then GetCaption = DDDefCurrent.Specification.Detail(i).sCaption
        i = i + 1
    Loop

End Function

Private Sub PointAtCurrentDDInList()
        '<EhHeader>
        On Error GoTo PointAtCurrentDDInList_Err
        '</EhHeader>

        Dim i As Long
100     i = 0
    
102     Do Until i = UBound(DDDefs) + 1

104         If DDDefs(i).desc = listDatabases.Text Then DDDefCurrent = DDDefs(i)
108         i = i + 1
        Loop

        '<EhFooter>
        Exit Sub

PointAtCurrentDDInList_Err:
        MsgBox "DynamicDataModule.PointAtCurrentDDInList_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Sub cmdEditData_Click()
        '<EhHeader>
        On Error GoTo cmdEditData_Click_Err
        '</EhHeader>

        Dim bAbort As Boolean
        Dim i As Long
        dxDBInspector1.BackColor = vbWhite
        
        If Not DDTableCurrent.IsLinkedTable And Len(dxDBGrid1.Columns(0).value) < 1 Then
100         'If Not DDTableCurrent.IsLinkedTable And (mDDRS.EOF Or mDDRS.Bof) Then
102         MsgBox "You must select a record before you can edit", vbExclamation, "Edit record"
104         bAbort = True
        End If
    
106     If Not bAbort Then

108         bAbort = False
110         C1TTab1Tab2.Visible = False
112         cmdEditData.Enabled = False
114         cmdAddData.Enabled = False
116         C1TTab1Tab2.CurrTab = 1
118         ListDataElements.Enabled = False
120         listDatabases.Enabled = False
122         cmdCancel.Visible = True
124         lblOperation.caption = "  Editing data..."

128         If DDTableCurrent.AllowEdit Then
                eCurrentOperation = Editing
            Else
                eCurrentOperation = browsing
            End If
            
            mDDRSGuid = dxDBGrid1.Columns(0).value
    
            If Not mDDRS.EOF Or Not mDDRS.Bof Then
    
                mDDRS.MoveFirst
                mDDRS.Find "[GUID1] = '" & mDDRSGuid & "'"
    
            End If
    
132         CopyRSValues mDDRS, mDDRSManual
    
134         If DDTableCurrent.IsGEOTable Then
            
136             SetGeoRsToCurrentUID
138             cmdSpatial.Enabled = True
140             cmdSpatial.Visible = True
            
            End If

144         Call UpdateCaptionOfNextButton
146         cmdNext.Visible = True
148         C1TTab1Tab2.Visible = True
            dxDBInspector1.Visible = True
            
            If eCurrentOperation = browsing Then
                dxDBInspector1.Container.Enabled = False
            Else
                dxDBInspector1.Container.Enabled = True
            End If
            
            i = 1

            Do Until i = dxDBInspector1.Count
                SetAutoCalcField DDDefCurrent.Prefix & DDTableCurrent.TableName, dxDBInspector1.Rows(i).FieldName
                i = i + 1
            Loop
            
            UpdateValidationMask DDTableCurrent.TableName
            UpdateProgressBarForNewTable
            lblLabel1.Visible = False
            
        End If
   
        '<EhFooter>
        Exit Sub

cmdEditData_Click_Err:
        MsgBox "DynamicDataModule.cmdEditData_Click_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Sub cmdNext_Click()
        '<EhHeader>
        On Error GoTo cmdNext_Click_Err
        '</EhHeader>
        
        Dim lIndex As Long
        Dim i As Long
        Dim bDonePostEditor As Boolean
        DoEvents
        bDonePostEditor = False
100     dxDBInspector1.PostEditor
        bDonePostEditor = True
        
102     dxDBInspector1.Visible = False
104     C1TTab1Tab2.Visible = False
106     cmdNext.Visible = False
108     cmdCancel.Visible = False
110     lblDescription.caption = "please wait momentarily for the form to load..."
112     lblDescription.FontSize = 15

114     DoEvents
    
        If CheckValidation(DDTableCurrent.TableName) = False Then
        
            cmdCancel.Visible = True
            cmdNext.Visible = True
    
116     ElseIf cmdNext.caption = "Next" Then
    
118         If DDDefCurrent.Specification.Detail(DDDefCurrent.Specification.lCurrentRank - 1).bIsMasterTable Then
120             CopyRSValues mDDRSManual, mDDRS
            End If
    
122         DDDefCurrent.Specification.lCurrentRank = DDDefCurrent.Specification.lCurrentRank + 1
124         MoveToNextTable DDDefCurrent.Specification.Detail(DDDefCurrent.Specification.lCurrentRank - 1).sTableName
        
126         UpdateCaptionOfNextButton
128         UpdateContainerOfControls True
        
130         If DDTableCurrent.IsLinkedTable Then
             
132             If SafeMoveFirst(mDDRSLinked) And mDDRSLinked.RecordCount > 0 Then
134                 CopyRSValues mDDRSLinked, mDDRSLinkedManual
136                 mDDRSLinkedGuid = mDDRSLinked.Fields("GUID2").value
                Else
138                 mDDRSLinkedManual.AddNew
140                 mDDRSLinkedGuid = GUIDGen

142                 If FieldExists(mDDRSLinkedManual, "GUID1") And Not FieldExists(mDDRSLinkedManual, "GUID2") Then
144                     mDDRSLinkedManual.AddNew Array("GUID1"), Array(mDDRSGuid)
146                 ElseIf FieldExists(mDDRSLinkedManual, "GUID2") And Not FieldExists(mDDRSLinkedManual, "GUID1") Then
148                     mDDRSLinkedManual.AddNew Array("GUID2"), Array(mDDRSLinkedGuid)
150                 ElseIf FieldExists(mDDRSLinkedManual, "GUID1") And FieldExists(mDDRSLinkedManual, "GUID2") Then
152                     mDDRSLinkedManual.AddNew Array("GUID1", "GUID2"), Array(mDDRSGuid, mDDRSLinkedGuid)
                    Else
154                     mDDRSLinkedManual.AddNew
                    End If
                
                End If
            
            End If
            
156         cmdCancel.Visible = True
158         cmdNext.Visible = True

            i = 1
If dxDBInspector1.Count > 0 Then
            Do Until i = dxDBInspector1.Count
19999           SetAutoCalcField DDDefCurrent.Prefix & DDTableCurrent.TableName, dxDBInspector1.Rows(i).FieldName
                i = i + 1
            Loop
            
            End If
            
1762        UpdateValidationMask DDTableCurrent.TableName
18654       dxProgressBar1.DoStep
            lProgressBarStep = lProgressBarStep + 1
            lblProgress.caption = "Step " & lProgressBarStep & " of " & lMaxProgressBarSteps
            
        Else 'Clicked FINISH
        
160         If Not DDTableCurrent.IsLinkedTable Then CopyRSValues mDDRSManual, mDDRS
162         If Not eCurrentOperation = browsing Then SaveData
164         DDDefCurrent.Specification.lCurrentRank = 1
166         mDDRS.Filter = adFilterNone
168         ReturnToBrowseTab

170         If Not DDTableCurrent.IsDDTable Then MoveToNextTable DDDefCurrent.Specification.Detail(DDDefCurrent.Specification.lCurrentRank - 1).sTableName

172         'dxDBGrid1.Dataset.ADODataset.Requery
174         'If eCurrentOperation = Editing Then mDDRSGuid = dxDBGrid1.Columns(0).Value
            If Not DDTableCurrent.IsLinkedTable Then Call ListDataElements_Click

176         If Len(mDDRSGuid) > 0 And (Not mDDRS.EOF Or Not mDDRS.Bof) Then
178             mDDRS.MoveFirst
180             mDDRS.Find "[GUID1] = '" & mDDRSGuid & "'"
                dxDBGrid1.Dataset.Locate "GUID1", mDDRSGuid, False, False
182             CopyRSValues mDDRS, mDDRSManual
            End If
            
184         UpdateContainerOfControls False
            lblLabel1.Visible = True
            dxProgressBar1.Visible = False
            lblProgress.Visible = False
            
        End If
    
186     If DDDefCurrent.Specification.lCurrentRank > 1 Then
188         cmdCancel.caption = "Back"
        Else
190         cmdCancel.caption = "Cancel"
        End If

192     UpdateDescription
194     dxDBInspector1.Visible = True
196     C1TTab1Tab2.Visible = True
        '<EhFooter>
        Exit Sub

cmdNext_Click_Err:
        If bDonePostEditor Then
            MsgBox "DynamicDataModule.cmdNext_Click_Err (" & Erl & ") " & Err.Description
        Else
       
            If dxDBInspector1.FocusedNode.Index > 0 Then
            Dim sFieldName As String
            sFieldName = dxDBInspector1.FocusedNode.Row.FieldName
            dxDBInspector1.Rows(0).Node.Focused = True
            dxDBInspector1.RowByName(sFieldName).Node.Focused = True
            MsgBox "Row value invlaid.  This data has been reset"
            End If
            
        End If
        
        dxDBInspector1.Visible = True
        cmdNext.Visible = True
        cmdCancel.Visible = True
        Exit Sub
        
        Resume Next
        '</EhFooter>
End Sub
    
Private Sub SetAutoCalcField(sTableName As String, _
                             sFieldName As String)
        '<EhHeader>
        On Error GoTo SetAutoCalcField_Err
        '</EhHeader>
                                      
        Dim i As Long
100     i = UBound(DDAutoCalcFields)
    
102     Do Until i < 0

104         If DDAutoCalcFields(i).sTableName = sTableName And DDAutoCalcFields(i).sFieldNameChild = sFieldName Then
            
106             If Not dxDBInspector1.RowByName(sFieldName) Is Nothing Then

108                 dxDBInspector1.RowByName(sFieldName).value = mDDRS.Fields(DDAutoCalcFields(i).sFieldNameParent).value
110                 dxDBInspector1.RowByName(sFieldName).ReadOnly = True
112                 dxDBInspector1.RowByName(sFieldName).LookUpRow.ListFieldName = ""
114                 dxDBInspector1.RowByName(sFieldName).Alignment = taLeftJustify ' taRightJustify
116                 dxDBInspector1_OnChangeNode dxDBInspector1.RowByName(sFieldName).Node, dxDBInspector1.RowByName(sFieldName).Node
            
                End If
        
            End If

118         i = i - 1
        Loop

        '<EhFooter>
        Exit Sub

SetAutoCalcField_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.DynamicDataModule.SetAutoCalcField", _
                  "DynamicDataModule component failure"
        '</EhFooter>
End Sub

Private Sub UpdateContainerOfControls(bGotoPage As Boolean)
        '<EhHeader>
        On Error GoTo UpdateContainerOfControls_Err
        '</EhHeader>
    
100     If DDTableCurrent.IsLinkedTable Then
    
102         SetParent dxDBInspector1.hwnd, C1ElasticLinkedTop.hwnd
104         Set dxDBInspector1.Container = C1ElasticLinkedTop
106         dxDBInspector1.Height = C1ElasticLinkedTop.Height - C1ElasticLinkedTop.BorderWidth - C1ElasticLinkedTop.BorderWidth
108         dxDBInspector1.Width = C1ElasticLinkedTop.Width - C1ElasticLinkedTop.BorderWidth - C1ElasticLinkedTop.BorderWidth
            
110         SetParent dxDBGrid1.ex.hwnd, C1ElasticLinkedBottom.hwnd
112         Set dxDBGrid1.Container = C1ElasticLinkedBottom
114         dxDBGrid1.Height = C1ElasticLinkedBottom.Height - C1ElasticLinkedBottom.BorderWidth - C1ElasticLinkedBottom.BorderWidth
116         dxDBGrid1.Width = C1ElasticLinkedBottom.Width - C1ElasticLinkedBottom.BorderWidth - C1ElasticLinkedBottom.BorderWidth
            
118         If bGotoPage Then
120             If Not C1TTab1Tab2.CurrTab = 0 Then C1TTab1Tab2.CurrTab = 2
            End If
        
        Else
            
122         SetParent dxDBInspector1.hwnd, C1ElasticSingle.hwnd
124         Set dxDBInspector1.Container = C1ElasticSingle
126         dxDBInspector1.Height = C1ElasticSingle.Height - C1ElasticSingle.BorderWidth - C1ElasticSingle.BorderWidth
128         dxDBInspector1.Width = C1ElasticSingle.Width - C1ElasticSingle.BorderWidth - C1ElasticSingle.BorderWidth

130         SetParent dxDBGrid1.ex.hwnd, C1ElasticBrowsing.hwnd
132         Set dxDBGrid1.Container = C1ElasticBrowsing
134         dxDBGrid1.Height = C1ElasticBrowsing.Height - C1ElasticBrowsing.BorderWidth - C1ElasticBrowsing.BorderWidth
136         dxDBGrid1.Width = C1ElasticBrowsing.Width - C1ElasticBrowsing.BorderWidth - C1ElasticBrowsing.BorderWidth

138         If bGotoPage Then
140             If Not C1TTab1Tab2.CurrTab = 0 Then C1TTab1Tab2.CurrTab = 1
            End If

        End If

        '<EhFooter>
        Exit Sub

UpdateContainerOfControls_Err:
        MsgBox "DynamicDataModule.UpdateContainerOfControls_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Sub UpdateCaptionOfNextButton()

    If Not DDTableCurrent.IsDDTable And (Not DDDefCurrent.Specification.lCurrentRank = DDDefCurrent.Specification.lMaxRank) Then
        cmdNext.caption = "Next"
    Else
        cmdNext.caption = "Save"
    End If

End Sub

Private Sub CopyRSValues(SourceRS As ADODB.Recordset, _
                         DestRS As ADODB.Recordset, _
                         Optional bCheckForAutoField As Boolean = False)
        '<EhHeader>
        On Error GoTo CopyRSValues_Err
        '</EhHeader>

        On Error Resume Next
        Dim i As Long

100     i = 0

102     If Not SourceRS Is Nothing Then

104         If Not SourceRS.EOF And Not SourceRS.Bof Then

106             Do Until i >= SourceRS.Fields.Count

108                 With DestRS.Fields(SourceRS.Fields(i).Name)
   
110                     If FieldExists(DestRS, .Name) Then

112                         .UnderlyingValue = Empty

114                         If SourceRS.Fields(SourceRS.Fields(i).Name).Type = adDate Then
116                             .value = Format(SourceRS.Fields(SourceRS.Fields(i).Name).value, "dd-MMM-yy")
                            Else
118                             .value = SourceRS.Fields(SourceRS.Fields(i).Name).value
                            End If

                        End If
                        
120                     If bCheckForAutoField Then
                        
122                         If Not dxDBInspector1.RowByName(.Name) Is Nothing Then
                            
124                             SetAutoCalcField DDDefCurrent.Prefix & DDTableCurrent.TableName, .Name
                            
                            End If
                        
                        End If

                    End With

126                 i = i + 1
                Loop

            End If

        End If

        '<EhFooter>
        Exit Sub

CopyRSValues_Err:
        MsgBox "DynamicDataModule.CopyRSValues_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Sub SaveData()
        '<EhHeader>
        On Error GoTo SaveData_Err
        '</EhHeader>

        Dim sDDTableNames() As String
        Dim sTableGeo As String
        Dim bEdit As Boolean
        Dim i As Long
        Dim bAbort As Boolean
        Dim sNewGUID As String
        Dim oShape As New TatukGIS_XDK10.XGIS_Shape
        Dim j As Long
        Dim sWKT As String
        Dim sDate As String
    
100     If DDTableCurrent.IsGEOTable Then
            
102         If mShape.IsEmpty Then
104             Set oShape = New TatukGIS_XDK10.XGIS_Shape
106         ElseIf mLayerCopy Is Nothing Then
108             Set oShape = mShape.CreateCopy
            Else
110             Set oShape = mLayerCopy.GetShape(mLayerCopy.GetLastUid)
            End If
        
112         If oShape.GetNumPoints = 0 Then
114             MsgBox "No spatial information has been saved for this record!  Save aborted", vbCritical
116             bAbort = True
            End If
        
        End If
        
118     If Not bAbort Then

            'Prepare form
120         lblDescription.caption = "please wait momentarily for the data to save............"
122         lblDescription.FontSize = 15
124         lblDescription.Refresh
126         UserControl.SetFocus
128         cmdCancel.Visible = False
130         lblOperation.caption = "  Browsing records..."

132         If DDTableCurrent.IsDDTable Then
            
                'Save edited dd table
134             If eCurrentOperation = Editing Then
136                 SynchHistoryEdit GUIDGen, mDDRSGuid, "DD ddDef Edit", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & DDTableCurrent.TableName, DDTableCurrent.IsGEOTable, "false", DDDefCurrent.Prefix
                Else
138                 SynchHistoryAddNew GUIDGen, mDDRSGuid, "DD ddDef Add", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & DDTableCurrent.TableName, DDTableCurrent.IsGEOTable, "false", DDDefCurrent.Prefix
                End If
                
140             j = 0

142             Do Until j = mDDRS.Fields.Count

144                 If mDDRS.Fields(j).Type = adDate Then
                        sDate = Format(mDDRS.Fields(j).value, "yyyymmdd")

146                     If sDate = "19991230" Or sDate = "18991230" Then
148                         mDDRS.Fields(j).value = Null
                        End If
                        
                    End If
                    
150                 j = j + 1
                Loop
                
152             If Not DDTableCurrent.IsGEOTable Then

                    If FieldExists(mDDRS, "WKT") Then
                        sWKT = "POINT (" & mDDRSManual.Fields("Longitude").value & " " & mDDRSManual.Fields("Latitude").value & ")"

                        If FieldExists(mDDRSManual, "WKT") Then mDDRSManual.Fields("WKT").value = sWKT
                        mDDRS.Fields("WKT").value = sWKT
                    End If
                    
154                 mDDRS.UpdateBatch adAffectCurrent

                Else
                    
156                 mLayer.SaveData
158                 mConn.Execute "DELETE FROM [" & DDDefCurrent.Prefix & DDTableCurrent.TableName & "] WHERE uid = " & mDDRS.Fields("UID").value
160                 mConn.Execute "DELETE FROM [" & DDDefCurrent.Prefix & left(DDTableCurrent.TableName, Len(DDTableCurrent.TableName) - 4) & "_GEO" & "] WHERE uid = " & mDDRS.Fields("UID").value
162                 Set oShape = mLayer.AddShape(oShape, True)
                    
164                 i = 0

166                 Do Until i = mDDRS.Fields.Count
                        
168                     With mDDRS.Fields(i)

170                         If Not .Name = "UID" Then
172                             oShape.SetField .Name, .value
                            End If

                        End With

174                     i = i + 1
                    Loop
                        
176                 If Len(sKMLString) > 1 And FieldExists(mDDRS, "KMLString") Then
178                     mDDRS.Fields("KMLString").value = sKMLString

180                     If FieldExists(mDDRSManual, "UID") Then mDDRSManual.Fields("KMLString").value = sKMLString
182                     oShape.SetField "KMLString", sKMLString
                    End If
                    
184                 If DDTableCurrent.IsGEOTable And FieldExists(mDDRS, "WKT") Then
186                     sWKT = oShape.ExportToWKT

188                     If FieldExists(mDDRSManual, "WKT") Then mDDRSManual.Fields("WKT").value = sWKT
190                     mDDRS.Fields("WKT").value = sWKT
192                     oShape.SetField "WKT", sWKT
                    End If
                    
194                 If Not DDTableCurrent.IsGEOTable And FieldExists(mDDRS, "Longitude") And FieldExists(mDDRS, "Latitude") Then
196                     sWKT = "POINT (" & mDDRSManual.Fields("Longitude").value & " " & mDDRSManual.Fields("Latitude").value & ")"

198                     If FieldExists(mDDRSManual, "WKT") Then mDDRSManual.Fields("WKT").value = sWKT
200                     mDDRS.Fields("WKT").value = sWKT
                    End If
                        
202                 mLayer.SaveData
204                 mDDRS.Fields("UID").value = oShape.uID
                
                End If

            Else
                
                'Save linked tables
206             i = 0

208             If DDDefCurrent.NumOfLinkedTables > 0 Then
              
210                 Do Until i >= UBound(mDDRSLinkedCollection)
    
212                     mDDRSLinkedCollection(i).RS.Filter = adFilterPendingRecords ' adFilterAffectedRecords

214                     If Not mDDRSLinkedCollection(i).RS.EOF Or Not mDDRSLinkedCollection(i).RS.Bof Then mDDRSLinkedCollection(i).RS.MoveFirst
                   
216                     Do Until mDDRSLinkedCollection(i).RS.EOF
                    
218                         If mDDRSLinkedCollection(i).RS.EditMode = adEditInProgress Then
220                             SynchHistoryEdit GUIDGen, mDDRSLinkedCollection(i).RS.Fields("GUID2").value, "DD linkedtable Edit", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & mDDRSLinkedCollection(i).sTableName, False, "false", DDDefCurrent.Prefix
222                         ElseIf mDDRSLinkedCollection(i).RS.EditMode = adEditAdd Then
224                             SynchHistoryAddNew GUIDGen, mDDRSLinkedCollection(i).RS.Fields("GUID2").value, "DD linkedtable Add", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & mDDRSLinkedCollection(i).sTableName, False, "false", DDDefCurrent.Prefix
226                         ElseIf mDDRSLinkedCollection(i).RS.EditMode = adEditDelete Then
228                             SynchHistoryDelete GUIDGen, mDDRSLinkedCollection(i).RS.Fields("GUID2").OriginalValue, "DD linkedtable Delete", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & mDDRSLinkedCollection(i).sTableName, False, "false", DDDefCurrent.Prefix
                            End If
    
230                         If FieldExists(mDDRSLinkedCollection(i).RS, "WKT") And FieldExists(mDDRSLinkedCollection(i).RS, "Longitude") And FieldExists(mDDRSLinkedCollection(i).RS, "Latitude") Then
232                             sWKT = "POINT (" & mDDRSLinkedCollection(i).RS.Fields("Longitude").value & " " & mDDRSLinkedCollection(i).RS.Fields("Latitude").value & ")"
234                             mDDRSLinkedCollection(i).RS.Fields("WKT").value = sWKT
                            End If
                            
236                         j = 0

238                         Do Until j = mDDRSLinkedCollection(i).RS.Fields.Count

240                             If mDDRSLinkedCollection(i).RS.Fields(j).Type = adDate Then

                                    sDate = Format(mDDRSLinkedCollection(i).RS.Fields(j).value, "yyyymmdd")

                                    If sDate = "19991230" Or sDate = "18991230" Then
244                                     mDDRSLinkedCollection(i).RS.Fields(j).value = Null
                                    End If
                        
                                End If
                    
246                             j = j + 1
                            Loop
                    
248                         mDDRSLinkedCollection(i).RS.MoveNext
                        Loop
                    
250                     mDDRSLinkedCollection(i).RS.UpdateBatch adAffectAll  ' adAffectAllChapters
252                     mDDRSLinkedCollection(i).RS.Filter = adFilterNone
254                     i = i + 1

                    Loop

                End If
                
256             mDDRS.Filter = adFilterPendingRecords
                
258             If Not mDDRS.EOF Then
                    
                    'Save mastertable
260                 If eCurrentOperation = Editing Then
262                     SynchHistoryEdit GUIDGen, mDDRSGuid, "DD mastertable Edit", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & "mastertable", False, "false", DDDefCurrent.Prefix
                    Else
264                     SynchHistoryAddNew GUIDGen, mDDRSGuid, "DD mastertable Add", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & "mastertable", False, "false", DDDefCurrent.Prefix
                    End If
                    
266                 j = 0

268                 Do Until j = mDDRS.Fields.Count

270                     If mDDRS.Fields(j).Type = adDate Then
                            sDate = Format(mDDRS.Fields(j).value, "yyyymmdd")

                            If sDate = "19991230" Or sDate = "18991230" Then
274                             mDDRS.Fields(j).value = Null
                            End If
                        
                        End If
                    
276                     j = j + 1
                    Loop
                
278                 If Not DDTableCurrent.IsGEOTable And FieldExists(mDDRS, "Longitude") And FieldExists(mDDRS, "Latitude") And FieldExists(mDDRS, "WKT") Then
280                     sWKT = "POINT (" & mDDRS.Fields("Longitude").value & " " & mDDRS.Fields("Latitude").value & ")"
282                     mDDRS.Fields("WKT").value = sWKT
                    End If
                    
284                 mDDRS.UpdateBatch 'adAffectCurrent
                
                End If

286             mDDRS.Filter = adFilterNone

            End If

288         Call SetAccessRights
290         cmdSpatial.Visible = False

292         If DDTableCurrent.IsGEOTable Then SleepAPI (4000)
294         'MsgBox "Saved"
296         Set mLayerCopy = Nothing

        End If
    
        '<EhFooter>
        Exit Sub

SaveData_Err:
        MsgBox "DynamicDataModule.SaveData_Err (" & Erl & " ) " & Err.Description
        Exit Sub
        Resume Next
        'Stop
        '</EhFooter>
End Sub

Function FieldExists(rstTest As ADODB.Recordset, ByVal fldName As String) As Boolean

    Dim fld As ADODB.Field
    FieldExists = False
    
    For Each fld In rstTest.Fields

        If fld.Name = right(fldName, Len(fld.Name)) Then
            FieldExists = True
            Exit For
        End If

    Next

    Set fld = Nothing

End Function

Public Sub SetGeoRsToCurrentUID()
        '<EhHeader>
        On Error GoTo SetGeoRsToCurrentUID_Err
        '</EhHeader>
        
100     If Not bLoadOfmDDRSInProgess Then
        
102         If Not mDDRS.Bof And Not mDDRS.EOF Then
    
104             Set mShape = mLayer.FindFirst(mLayer.Extent, "GIS_UID = " & mDDRS.Fields("UID").value, Nothing, "", True)

106             If mShape Is Nothing Then
108                 MsgBox "mShape not found!", vbCritical
                    'Stop
                End If
            End If
        
        End If
        
        Dim utils As New TatukGIS_XDK10.XGIS_Utils

        '<EhFooter>
        Exit Sub

SetGeoRsToCurrentUID_Err:
        MsgBox "DynamicDataModule.SetGeoRsToCurrentUID_Err (" & Erl & " ) " & Err.Description
        
        '</EhFooter>
End Sub

Private Sub cmdSpatial_Click()
        '<EhHeader>
        On Error GoTo cmdSpatial_Click_Err
        '</EhHeader>
        
        dxDBInspector1.PostEditor

100     If mLayerCopy Is Nothing Then
102         Set mLayerCopy = New TatukGIS_XDK10.XGIS_LayerVector
104         mLayerCopy.caption = "Temp layer for new shape info"
        
106         If Not mShape Is Nothing Then
        
108             Set mShapeNew = mLayerCopy.AddShape(mShape.CreateCopy, True)
110             mShapeNew.CopyFields mShape
        
            Else
        
112             Set mShapeNew = mLayerCopy.CreateShape(XgisShapeTypePoint) ' AddShape(mShape.CreateCopy, True)
114             Set mShape = mShapeNew
        
            End If
        End If
    
116     dxDBInspector1.SetFocus
118     RaiseEvent GetSpatialLoc(mLayerCopy)

        '<EhFooter>
        Exit Sub

cmdSpatial_Click_Err:
        MsgBox "DynamicDataModule.cmdSpatial_Click_Err (" & Erl & " ) " & Err.Description
        
        '</EhFooter>
End Sub

Private Sub cmdUpdateLinked_Click()
    dxDBInspector1.PostEditor

    CopyRSValues mDDRSLinkedManual, mDDRSLinked

    If DDTableCurrent.TableName = "linkItemsDisAndRec" Then NomadTotalPriceCalc mDDRSLinked
End Sub

Private Sub CopyToClipboard_Click()

    dxDBGrid1.M.SelectAll
    dxDBGrid1.M.CopySelectedToClipboard
    MsgBox "The contents of the grid have been copied to the clipboard", vbInformation, "Clipboard"
    dxDBGrid1.M.ClearSelection
    
End Sub

Private Sub dxDBGrid1_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, _
                                   ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
        '<EhHeader>
        On Error GoTo dxDBGrid1_OnChangeNode_Err
        '</EhHeader>

100     If Not OldNode Is Nothing Then

            ' If Not OldNode.RecNo = Node.RecNo Then
        
102         If Not DDTableCurrent.IsLinkedTable Then
            
104             If Not mDDRS.EOF Then
106                 mDDRS.MoveFirst
108                 mDDRS.Find "[GUID1] = '" & dxDBGrid1.Columns(0).value & "'"
110                 CopyRSValues mDDRS, mDDRSManual
                End If

            Else

112             If Not mDDRSLinked.EOF Then
                
114                 mDDRSLinkedGuid = mDDRSLinked.Fields("GUID2").value
116                 CopyRSValues mDDRSLinked, mDDRSLinkedManual
                    cmdUpdateLinked.Enabled = True
                End If

            End If

            '  End If
        End If

        'Resume Next
        '<EhFooter>
        Exit Sub

dxDBGrid1_OnChangeNode_Err:
        MsgBox "DynamicDataModule.dxDBGrid1_OnChangeNode_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Sub dxDBGrid1_OnDblClick()
        '<EhHeader>
        On Error GoTo dxDBGrid1_OnDblClick_Err
        '</EhHeader>
        
        Dim sFilter As String
        Dim sTitle As String
        Dim i As Long
        Dim sTableNames() As String
        Dim bPreventDeletion As Boolean
        Dim oRS As ADODB.Recordset
        
100     i = 0
102     bPreventDeletion = True

104     If DDTableCurrent.IsLinkedTable Then
106         Set oRS = mDDRSLinked
        Else
108         Set oRS = mDDRS
        End If

110     If DDTableCurrent.AllowDelete Then

112         If Not oRS.EOF And Not oRS.Bof And Not oRS.RecordCount = 0 Then
    
114             If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Confirm deletion") = vbYes Then
                
116                 If DDTableCurrent.IsDDTable Then
                    
118                     bPreventDeletion = CheckIfDDWithFieldValueExists(oRS.Fields(0).value, DDDefCurrent.Prefix & DDTableCurrent.TableName)
                    
120                     If bPreventDeletion Then
122                         DebugPrint "Deletion is not possible since this value is used in at least one other table"
124                         MsgBox "Deletion is not possible since this record is used elsewhere in the database and deleting this record would corrupt that information"
                        End If

126                 ElseIf DDTableCurrent.IsMasterTable Then

128                     bPreventDeletion = False
                    
130                     If DDDefCurrent.NumOfLinkedTables > 0 Then
                    
132                         Do Until bPreventDeletion Or i = UBound(mDDRSLinkedCollection) + 1
    
134                             If Not mDDRSLinkedCollection(i).sTableName = "" Then
136                                 bPreventDeletion = CheckIfFieldWithValueExists(oRS.Fields(0).value, "GUID1", DDDefCurrent.Prefix & mDDRSLinkedCollection(i).sTableName)

138                                 If bPreventDeletion Then MsgBox ("Deletion is not possible since this record is used elsewhere in the database and deleting this record would corrupt that information.  (Try clicking 'Next' to see if there are any records logged)")
                                End If
    
140                             i = i + 1
                            Loop
                        
                        End If
                    
142                 ElseIf DDTableCurrent.IsLinkedTable Then
                    
144                     bPreventDeletion = False
                        
                    End If
                
146                 If Not bPreventDeletion Then

148                     If Not DDTableCurrent.IsLinkedTable Then
150                         SynchHistoryDelete GUIDGen, oRS.Fields("GUID1").value, "OASIS DD Module Delete", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & DDTableCurrent.TableName, False, "false", DDDefCurrent.Prefix
                        Else
152                         SynchHistoryDelete GUIDGen, oRS.Fields("GUID2").value, "OASIS DD Module Delete", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & DDTableCurrent.TableName, False, "false", DDDefCurrent.Prefix
                        End If
                        
154                     If DDTableCurrent.IsGEOTable Then
156                         mConn.Execute "DELETE FROM [" & DDDefCurrent.Prefix & DDTableCurrent.TableName & "] WHERE [UID] = " & oRS.Fields("UID").value
158                         mConn.Execute "DELETE FROM [" & DDDefCurrent.Prefix & Replace(DDTableCurrent.TableName, "_FEA", "") & "_GEO] WHERE [UID] = " & oRS.Fields("UID").value
160                     ElseIf DDTableCurrent.IsDDTable Then
162                         oRS.Delete adAffectCurrent
164                     ElseIf DDTableCurrent.IsLinkedTable Then
166                         dxDBGrid1.Dataset.Delete
                        Else
                        
168                         mConn.Execute "DELETE FROM [" & DDDefCurrent.Prefix & DDTableCurrent.TableName & "] WHERE [GUID1] = '" & oRS.Fields("GUID1").value & "'"
                        End If
                            
170                     If Not DDTableCurrent.IsLinkedTable Then
                            
172                         oRS.UpdateBatch adAffectCurrent
174                         mDDRSManual.AddNew

176                         If Not bSQLServerInUse Then dxDBGrid1.Dataset.ADODataset.Requery
178                         mDDRSGuid = dxDBGrid1.Columns(0).value
                            
180                         If Not mDDRS.EOF Then
182                             mDDRS.MoveFirst
184                             mDDRS.Find "[GUID1] = '" & mDDRSGuid & "'"
186                             CopyRSValues mDDRS, mDDRSManual
                            End If
                            
                        Else
                            'If SafeMoveFirst(mDDRSLinked) Then CopyRSValues mDDRSLinked, mDDRSLinkedManual
                           
188                         If SafeMoveFirst(mDDRSLinked) And mDDRSLinked.RecordCount > 0 Then
190                             CopyRSValues mDDRSLinked, mDDRSLinkedManual
192                             mDDRSLinkedGuid = mDDRSLinked.Fields("GUID2").value
                            Else
194                             mDDRSLinkedManual.AddNew
196                             mDDRSLinkedGuid = GUIDGen

198                             If FieldExists(mDDRSLinkedManual, "GUID1") And Not FieldExists(mDDRSLinkedManual, "GUID2") Then
200                                 mDDRSLinkedManual.AddNew Array("GUID1"), Array(mDDRSGuid)
202                             ElseIf FieldExists(mDDRSLinkedManual, "GUID2") And Not FieldExists(mDDRSLinkedManual, "GUID1") Then
204                                 mDDRSLinkedManual.AddNew Array("GUID2"), Array(mDDRSLinkedGuid)
206                             ElseIf FieldExists(mDDRSLinkedManual, "GUID1") And FieldExists(mDDRSLinkedManual, "GUID2") Then
208                                 mDDRSLinkedManual.AddNew Array("GUID1", "GUID2"), Array(mDDRSGuid, mDDRSLinkedGuid)
                                Else
210                                 mDDRSLinkedManual.AddNew
                                End If
                
                            End If
                           
                        End If
                            
                    End If

                End If
                
            End If

        End If

        '<EhFooter>
        Exit Sub

dxDBGrid1_OnDblClick_Err:
        MsgBox "DynamicDataModule.dxDBGrid1_OnDblClick_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Function RecordSetHasField(strField As String, _
                                   RS As ADODB.Recordset) As Boolean
    Dim objField As Field

    For Each objField In RS.Fields

        If UCase(objField.Name) = UCase(strField) Then
            RecordSetHasField = True
            Exit Function
        End If

    Next

    RecordSetHasField = False
End Function

Private Function CheckIfFieldWithValueExists(sValue As String, _
                                             sFieldName As String, _
                                             sTableName As String) As Boolean
        '<EhHeader>
        On Error GoTo CheckIfFieldWithValueExists_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
100     oRS.Open "SELECT [" & sFieldName & "] FROM [" & sTableName & "] WHERE [" & sFieldName & "] = '" & sValue & "'", mConn, adOpenDynamic, adLockReadOnly

102     If oRS.RecordCount > 0 Then
104         CheckIfFieldWithValueExists = True
        Else
106         CheckIfFieldWithValueExists = False
        End If
    
108     oRS.Close
110     Set oRS = Nothing

        '<EhFooter>
        Exit Function

CheckIfFieldWithValueExists_Err:
        MsgBox "DynamicDataModule.CheckIfFieldWithValueExists_Err (" & Erl & " ) " & Err.Description
        
        '</EhFooter>
End Function

Private Function CheckIfDDWithFieldValueExists(sValue As String, _
                                               sFieldName As String) As Boolean
        '<EhHeader>
        On Error GoTo CheckIfDDWithFieldValueExists_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
        Dim oRSTables As New ADODB.Recordset
        Dim oRSColumns As New ADODB.Recordset
 
100     With oRSTables
        
102         .Open mConn.OpenSchema(adSchemaTables) ', , adOpenDynamic, adLockBatchOptimistic
104         .Filter = "TABLE_NAME LIKE 'dd_%' AND TABLE_TYPE = 'TABLE'"

106         If Not .EOF Or Not .Bof Then .MoveFirst
            oRSColumns.Open mConn.OpenSchema(adSchemaColumns)
            
108         Do Until .EOF

                oRSColumns.Filter = "COLUMN_NAME = '" & sFieldName & "' AND TABLE_NAME = '" & oRSTables.Fields("TABLE_NAME").value & "'"
110
112             oRSColumns.Filter = "COLUMN_NAME = '" & sFieldName & "' AND TABLE_NAME = '" & oRSTables.Fields("TABLE_NAME").value & "'"

114             If Not oRSColumns.EOF Or Not oRSColumns.Bof Then oRSColumns.MoveFirst
            
116             Do Until oRSColumns.EOF
                
118                 oRS.Open "SELECT top 1 [" & sFieldName & "] FROM [" & oRSTables.Fields("TABLE_NAME").value & "] WHERE [" & sFieldName & "] = '" & sValue & "'", mConn, adOpenDynamic, adLockReadOnly
    
120                 If oRS.RecordCount > 0 Then
122                     CheckIfDDWithFieldValueExists = True
124                     Set oRS = Nothing
126                     Set oRSTables = Nothing
128                     Set oRSColumns = Nothing
                        Exit Function
                    Else
130                     CheckIfDDWithFieldValueExists = False
                    End If
                    
                    oRS.Close
132                 oRSColumns.MoveNext
                Loop
                
134
            
136             .MoveNext
            Loop
            
            Set oRSColumns = Nothing
        
138         Set oRSTables = Nothing
        
        End With
    
        'oRS.Close
140     Set oRS = Nothing

        '<EhFooter>
        Exit Function

CheckIfDDWithFieldValueExists_Err:
        MsgBox "DynamicDataModule.CheckIfDDWithFieldValueExists_Err (" & Erl & " ) " & Err.Description
        
        '</EhFooter>
End Function

Private Sub dxDBGrid1_OnMouseDown(ByVal Button As Long, _
                                  ByVal Shift As Long, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)

    If Button = 2 Then PopupMenu ExportOptions

End Sub

Public Sub Init(ByRef List0 As ListBox, _
                ByRef List1 As ListBox)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        
        Dim i As Long
        Dim l As Long
        Dim iDDDefIndex As Long
        Dim oRS As New ADODB.Recordset
        Dim oConn As ADODB.Connection
        Dim bNotAvailable As Boolean
        Dim sConnectionString As String
        Dim sExcludedFields As String
        Dim sLockedFields As String
        
100     SavedIdAssigned = False
        
102     'Set C1Elastic666.Picture = C1ElasticSimple.Picture
104     'Set C1Elastic667.Picture = C1ElasticSimple.Picture
106     'Set C1Elastic1.Picture = C1ElasticSimple.Picture
        
108     lblDescription.caption = "Welcome to the Data Entry Module.  Please select a database on the left to begin......"
110     lblDescription.FontSize = 15
112     C1TTab1Tab2.TabHeight = 1

114     Set listDatabases = List0
116     listDatabases.Clear
118     Set ListDataElements = List1
120     ListDataElements.Clear
        
122     Set dxDBGrid1.DataSource = Nothing
124     dxDBGrid1.Columns.DestroyColumns
126     C1ElasticSimple.Refresh
128     Set mDDRS = New ADODB.Recordset

        If bSQLServerInUse Then
129         oRS.Open "SELECT * FROM DynamicDataDefs WHERE [EnableDataEntry] = 'TRUE' ORDER BY cast([Description] as nvarchar(255))", m_Cnn, adOpenDynamic, adLockReadOnly
        Else
130         oRS.Open "SELECT * FROM DynamicDataDefs WHERE [EnableDataEntry] = TRUE ORDER BY [Description]", m_Cnn, adOpenDynamic, adLockReadOnly
        End If
    
132     ReDim DDRecordSets(0)
134     ReDim DDDefs(0)
136     ReDim AllDDDefTables(0)
    
138     l = 0
140     iDDDefIndex = 0
    
142     Do Until oRS.EOF

144         bNotAvailable = False
146         Set oConn = New ADODB.Connection

            If bSQLServerInUse Then
        
                oConn.Open g_sGlobalConnectionString
            Else
        
148             oConn.ConnectionString = DDDefCurrent.ConnectionString
150             sConnectionString = Replace(oRS.Fields("ConnectionString").value, "\data\db\dynamicdata", AppPath & "data\db\dynamicdata", , , vbTextCompare)
152             oConn.ConnectionString = sConnectionString
154             oConn.CursorLocation = g_sGlobalCursorLocation
                On Error GoTo NotAvailable
156             oConn.Open
                On Error GoTo Init_Err
            
            End If

158         If oConn.State = adStateOpen Then oConn.Close
            
160         If Not bNotAvailable Then
    
162             iDDDefIndex = iDDDefIndex + 1
        
164             ReDim Preserve DDDefs(UBound(DDDefs) + 1)
166             DDDefs(UBound(DDDefs)).Name = oRS.Fields("DDDefName").value
168             DDDefs(UBound(DDDefs)).desc = oRS.Fields("Description").value
170             DDDefs(UBound(DDDefs)).ListIndex = iDDDefIndex - 1
172             DDDefs(UBound(DDDefs)).Prefix = "dd_" & oRS.Fields("DDDefName").value & "_"
174             DDDefs(UBound(DDDefs)).ConnectionString = oRS.Fields("ConnectionString").value

                If FieldExists(oRS, "ExcludedFields") Then sExcludedFields = IIf(Len(oRS.Fields("ExcludedFields").value) > 0, oRS.Fields("ExcludedFields").value, "")
                DDDefs(UBound(DDDefs)).ExcludedFields = sExcludedFields

                If FieldExists(oRS, "LockedFields") Then sLockedFields = IIf(Len(oRS.Fields("LockedFields").value) > 0, oRS.Fields("LockedFields").value, "")
                DDDefs(UBound(DDDefs)).LockedFields = sLockedFields
                
176             GetDDDetailedInfo DDDefs(UBound(DDDefs)).Prefix, IIf(IsNull(oRS.Fields("AccessRights").value), "", oRS.Fields("AccessRights").value), DDDefs(UBound(DDDefs))

178             If Len(oRS.Fields("AccessRights").value) > 0 Then
180                 GetDDDetailedSpec DDDefs(UBound(DDDefs)).Prefix, DDDefs(UBound(DDDefs)), sConnectionString
182                 listDatabases.AddItem oRS.Fields("Description").value ', iDDDefIndex - 1
                End If
            End If

184         oRS.MoveNext
        
        Loop
        
186     oRS.Close
188     Set oRS = Nothing
190     Set dxDBInspector1.DataSource = Nothing

        Exit Sub
NotAvailable:
        bNotAvailable = True
        Resume Next
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox "Your configuration settings for the [" & oRS.Fields("Description").value & "] module are in error.  Please contact an OASIS administrator", vbExclamation, "Configuration error"
        oRS.Close
        Set oRS = Nothing
        Set dxDBInspector1.DataSource = Nothing
        'MsgBox "DynamicDataModule.Init_Err (" & Erl & " ) " & Err.Description
        'Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    Set mDDRS = Nothing
    Set mDDRSLinked = Nothing
    Set mConn = Nothing
    Unload dxDBInspector1

End Sub

Private Sub ScanForDDFields()
        '<EhHeader>
        On Error GoTo ScanForDDFields_Err
        '</EhHeader>

        Dim i As Long
        Dim oRS As New ADODB.Recordset
    
100     If DDTableCurrent.IsLinkedTable Then
102         Set oRS = mDDRSLinked
        Else
104         Set oRS = mDDRS
        End If
    
106     i = 0

108     Do Until i = oRS.Fields.Count
    
110         If left(oRS.Fields(i).Name, 2) = "dd" Or left(oRS.Fields(i).Name, 11) = "mastertable" Or left(oRS.Fields(i).Name, 4) = "link" Then

112             If C1TTab1Tab2.CurrTab = 0 Or Not dxDBInspector1.RowByName(oRS.Fields(i).Name) Is Nothing Then
114                 SetDropdown oRS.Fields(i).Name ', IIf(C1TTab1Tab2.CurrTab = 0 Or DDTableCurrent.IsLinkedTable, True, False)
                End If

            End If

116         If InStr(LCase(oRS.Fields(i).Name), "qty") < 1 And oRS.Fields(i).Type = 5 Then
118             If Not dxDBInspector1.RowByName(oRS.Fields(i).Name) Is Nothing Then dxDBInspector1.RowByName(oRS.Fields(i).Name).DecimalPlaces = oRS.Fields(i).DefinedSize '15
            ElseIf oRS.Fields(i).Type = 5 Then

                If Not dxDBInspector1.RowByName(oRS.Fields(i).Name) Is Nothing Then dxDBInspector1.RowByName(oRS.Fields(i).Name).DecimalPlaces = oRS.Fields(i).DefinedSize '4
            End If
            
            If oRS.Fields(i).Name = "Longitude" Or oRS.Fields(i).Name = "Latitude" Or oRS.Fields(i).Name = "MGRS" Then
                If Not dxDBInspector1.RowByName(oRS.Fields(i).Name) Is Nothing Then
                    dxDBInspector1.RowByName(oRS.Fields(i).Name).RowType = iedButtonEdit
                    dxDBInspector1.RowByName(oRS.Fields(i).Name).ButtonRow.ButtonOnly = True
                    dxDBInspector1.RowByName(oRS.Fields(i).Name).DecimalPlaces = 15
                    'dxDBInspector1.RowByName(oRS.Fields(i).Name).ButtonRow.HideEditCursor = True
                End If
            
            End If

            If Not dxDBInspector1.RowByName(oRS.Fields(i).Name) Is Nothing Then
                If dxDBInspector1.RowByName(oRS.Fields(i).Name).RowType = iedDateEdit Then
                    dxDBInspector1.RowByName(oRS.Fields(i).Name).DateRow.ShowClearButton = True
                End If

                dxDBInspector1.RowByName(oRS.Fields(i).Name).Alignment = taLeftJustify
            
            End If

120         i = i + 1
        Loop

        '<EhFooter>
        Exit Sub

ScanForDDFields_Err:
        MsgBox "DynamicDataModule.ScanForDDFields_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Sub dxDBInspector1_OnChangeNode(ByVal OldNode As DXDBINSPLibCtl.IdxDBRowNode, _
                                        ByVal Node As DXDBINSPLibCtl.IdxDBRowNode)

    Dim sLookupTableName As String
        
    If left(right$(OldNode.Row.ObjectName, 6), 4) = "-PAD" Then
        sLookupTableName = left(OldNode.Row.ObjectName, Len(OldNode.Row.ObjectName) - 6)
    ElseIf left(right$(OldNode.Row.ObjectName, 5), 4) = "-PAD" Then
        sLookupTableName = left(OldNode.Row.ObjectName, Len(OldNode.Row.ObjectName) - 5)
    Else
        sLookupTableName = OldNode.Row.ObjectName
    End If
        
    If Not GetComboRelationChildField(DDDefCurrent.Prefix & DDTableCurrent.TableName, OldNode.Row.ObjectName) = "" Then
        SetDropdown GetComboRelationChildField(DDDefCurrent.Prefix & DDTableCurrent.TableName, OldNode.Row.ObjectName), " WHERE [" & sLookupTableName & "] = '" & OldNode.Row.value & "' "
    End If
    
    If CheckValidationByNode(OldNode, DDTableCurrent.TableName) = False Then
        OldNode.Focused = True
    End If
    
    'dude

End Sub

Private Sub dxDBInspector1_OnDateEditValidation(ByVal RowNode As DXDBINSPLibCtl.IdxDBRowNode, _
                                                ByVal AText As String, _
                                                ADate As Variant, _
                                                AErrorMessage As String, _
                                                AError As Boolean)
    'Stop
End Sub

Private Sub dxDBInspector1_OnDrawCaption(ByVal RowNode As DXDBINSPLibCtl.IdxDBRowNode, _
                                         ByVal hdc As Long, _
                                         ByVal l As Single, _
                                         ByVal T As Single, _
                                         ByVal R As Single, _
                                         ByVal b As Single, _
                                         AText As String, _
                                         ByVal AFont As DXDBINSPLibCtl.IFontDisp, _
                                         AFontColor As Long, _
                                         AColor As Long, _
                                         Done As Boolean)
    
    If RowNode.Row.ReadOnly Then
        AFontColor = vbBlack
        AFont.Bold = False
        'AFont.Size = 8
        AColor = RGB(186, 184, 184)
    Else
        AFont.Bold = True
    End If

End Sub

Private Sub dxDBInspector1_OnDrawValue(ByVal RowNode As DXDBINSPLibCtl.IdxDBRowNode, _
                                       ByVal hdc As Long, _
                                       ByVal l As Single, _
                                       ByVal T As Single, _
                                       ByVal R As Single, _
                                       ByVal b As Single, _
                                       AText As String, _
                                       ByVal AFont As DXDBINSPLibCtl.IFontDisp, _
                                       AFontColor As Long, _
                                       AColor As Long, _
                                       Done As Boolean)

    If RowNode.Row.ReadOnly Then
        AFontColor = RGB(143, 140, 140) ' vbHiragana
        AFont.Bold = True
        'AFont.Size = 8
        AColor = vbWhite 'vbMagenta
        
    Else
        AFontColor = vbBlack
        AFont.Bold = False
    End If

End Sub

Private Sub dxDBInspector1_OnEditButtonClick(ByVal RowNode As dxDBRowNode)
    
    Dim oPoint As New TatukGIS_XDK10.XGIS_Point
    Dim sMGRS As String
    Dim X As Double
    Dim Y As Double
              
    Select Case RowNode.Row.ObjectName
    
        Case "Longitude"

            If RowNode.Row.value = 0 Or IsNull(RowNode.Row.value) Or RowNode.Row.value = "" Then
                RaiseEvent GetCurrentExtentCentroid(oPoint)
            Else
                oPoint.Prepare dxDBInspector1.RowByName("Longitude").value, dxDBInspector1.RowByName("Latitude").value
            End If
                
            RaiseEvent GetSpatialLocEx(oPoint)
                
        Case "Latitude"

            If RowNode.Row.value = 0 Or IsNull(RowNode.Row.value) Or RowNode.Row.value = "" Then
                RaiseEvent GetCurrentExtentCentroid(oPoint)
            Else
                oPoint.Prepare dxDBInspector1.RowByName("Longitude").value, dxDBInspector1.RowByName("Latitude").value
            End If
               
            RaiseEvent GetSpatialLocEx(oPoint)
                    
        Case "MGRS"

            If RowNode.Row.value = 0 Or IsNull(RowNode.Row.value) Or RowNode.Row.value = "" Then
                RaiseEvent GetCurrentExtentCentroid(oPoint)
            Else
                RaiseEvent ConvertMGRStoPT(dxDBInspector1.RowByName("MGRS").value, X, Y)
                oPoint.Prepare X, Y
            End If
               
            RaiseEvent GetSpatialLocEx(oPoint)
                    
    End Select
    
    Set oPoint = Nothing
    
End Sub

Private Sub SetGridCaptions()
        '<EhHeader>
        On Error GoTo SetGridCaptions_Err
        '</EhHeader>
    
        Dim i As Long
        Dim oRSCaption As ADODB.Recordset
        Dim oRS As New ADODB.Recordset
        Dim oRSManual As New ADODB.Recordset
        Dim sCaption As String
        Dim lCountOfFieldsNotMemo As Long
        Dim lCountOfFieldsLengthsNotMemo As Long
        
        Dim oDB As New ADOX.Catalog
        Dim itbl As New ADOX.Table
        Dim iCol As New ADOX.Column
        Dim j As Long
              
100     If DDTableCurrent.IsLinkedTable Then
102         Set oRS = mDDRSLinked
104         Set oRSManual = mDDRSLinkedManual
        Else
106         Set oRS = mDDRS
108         Set oRSManual = mDDRSManual
        End If
   
112     lCountOfFieldsLengthsNotMemo = 0
114     lCountOfFieldsNotMemo = 0
116     dxDBGrid1.Option = egoAutoWidth
118     dxDBGrid1.OptionEnabled = True
        
120     Set oDB.ActiveConnection = mConn
        
122     'For Each itbl In oDB.Tables

        '  Set itbl = oDB.Tables(DDTableCurrent.TableName)
    
        Do Until i = oDB.Tables.Count
    
            If oDB.Tables.Item(i).Name = DDDefCurrent.Prefix & DDTableCurrent.TableName Then Set itbl = oDB.Tables.Item(i)
            i = i + 1
        Loop
    
        i = 0
    
124     If 1 Or itbl.Name = oRS.Fields(0).Properties(1) Then
                
126         For Each iCol In itbl.Columns
                
                If bSQLServerInUse Then
                    Set oRSCaption = New ADODB.Recordset
                    oRSCaption.Open "SELECT * FROM ::fn_listExtendedProperty ( 'MS_Description','user', 'dbo', 'table', '" & itbl.Name & "', 'column', '" & iCol.Name & "')", mConn, adOpenDynamic, adLockBatchOptimistic

                    If Not oRSCaption.EOF Then sCaption = oRSCaption.Fields("value").value
                    Set oRSCaption = Nothing
                Else
                
128                 sCaption = iCol.Properties(2).value

                End If
                
130             If Not dxDBInspector1.RowByName(iCol.Name) Is Nothing Then
                
132                 With dxDBInspector1.RowByName(iCol.Name)
                    
134                     .Alignment = taLeftJustify
                    
136                     If iCol.Type = adDate Or iCol.Type = adNumeric Or iCol.Type = adBoolean Or iCol.Type = 1 Then
138                         lCountOfFieldsNotMemo = lCountOfFieldsNotMemo + 1
                        End If

140                     If .FieldName = "GUID1" Or .FieldName = "GUID2" Or .FieldName = "UID" Then .ReadOnly = True
142                     If Not sCaption = "" And Not IsNull(sCaption) Then .caption = sCaption

                    End With
                    
                End If
                    
144             j = 0

146             Do Until j = dxDBGrid1.Columns.Count
                            
148                 If dxDBGrid1.Columns(j).FieldName = iCol.Name Then
150                     If Len(sCaption) > 0 Then dxDBGrid1.Columns(j).caption = sCaption
                    End If
                        
                    If dxDBGrid1.Columns(j).FieldType = xftDateTime Then
                        dxDBGrid1.Columns(j).DisplayFormat = "Medium Date"
                    End If

156                 j = j + 1
                            
                Loop
                
            Next

            '   Exit For
        End If

        'Next
        
        If (C1TTab1Tab2.CurrTab = 0 Or DDTableCurrent.IsLinkedTable) Then
        
158         i = 0
        
160         Do Until i = dxDBGrid1.Dataset.FieldCount
                dxDBGrid1.Columns(i).Width = dxDBGrid1.Width / dxDBGrid1.Dataset.FieldCount
170             i = i + 1
            Loop
        
        End If
        
        Set oDB = Nothing
        Set itbl = Nothing
       
        '<EhFooter>
        Exit Sub

SetGridCaptions_Err:
        MsgBox "DynamicDataModule.SetGridCaptions_Err (" & Erl & " ) " & Err.Description
        
        '</EhFooter>
End Sub

'Private Function DoesTableExist(cn As ADODB.Connection, _
'                                sTable As String) As Boolean
'    Dim Rs As New ADODB.Recordset
'
'    On Error GoTo hell
'    Rs.Open "SELECT * FROM [" & sTable & "]", cn, adOpenForwardOnly, adLockReadOnly
'
'    DoesTableExist = True
'
'    On Error Resume Next
'    Rs.Close
'    Set Rs = Nothing
'
'    Exit Function
'hell:
'    On Error Resume Next
'    Err.Clear
'
'    DoesTableExist = False
'    Set Rs = Nothing
'End Function

Private Sub SetDropdown(sField As String, _
                        Optional sWhereClause As String)
        '<EhHeader>
        On Error GoTo SetDropdown_Err
        '</EhHeader>

        Dim oRS As ADODB.Recordset
        Dim sfieldOLD As String
        Dim sSQL As String
        Dim bDoesExist As Boolean
        Dim sLookupTableName As String
        
        If left(right$(sField, 6), 4) = "-PAD" Then
            sLookupTableName = left(sField, Len(sField) - 6)
        ElseIf left(right$(sField, 5), 4) = "-PAD" Then
            sLookupTableName = left(sField, Len(sField) - 5)
        Else
            sLookupTableName = sField
        End If
        
        sSQL = "SELECT DISTINCT [GUID1], [option] FROM [" & sLookupTableName & "] " & sWhereClause & " order by [option]"
100     sfieldOLD = sField
          
120     If Not dxDBInspector1.RowByName(sfieldOLD) Is Nothing Then

122         Set oRS = New ADODB.Recordset
124         oRS.Open sSQL, mConn, adOpenKeyset, adLockReadOnly ' ', adOpenStatic, adLockReadOnly
126         dxDBInspector1.RowByName(sfieldOLD).RowType = iedLookupEdit
128         dxDBInspector1.RowByName(sfieldOLD).LookUpRow.ListFieldName = "option"
130         dxDBInspector1.RowByName(sfieldOLD).LookUpRow.LookUpKeyFieldName = "GUID1"
132         dxDBInspector1.RowByName(sfieldOLD).LookUpRow.LookUpDisplayFieldName = "option"
134         dxDBInspector1.RowByName(sfieldOLD).LookUpRow.ListSearchIndex = 1
136         dxDBInspector1.RowByName(sfieldOLD).LookUpRow.ListColumns = "option" '"*"
138         Set dxDBInspector1.RowByName(sfieldOLD).LookUpRow.DB.Recordset = oRS

140         bDoesExist = False

142         Do Until oRS.EOF
            
144             If oRS.Fields(0).value = dxDBInspector1.RowByName(sfieldOLD).value Then
146                 bDoesExist = True
                End If
            
148             oRS.MoveNext
            Loop

150         If Not oRS.Bof Then oRS.MoveFirst

152         If Not bDoesExist Then
            
                'wipe the child
154             dxDBInspector1.RowByName(sfieldOLD).value = ""
            
156             If Not sWhereClause = "" And Not GetComboRelationChildField(DDDefCurrent.Prefix & DDTableCurrent.TableName, sfieldOLD) = "" Then
                    'child is a parent also!

158                 SetDropdown GetComboRelationChildField(DDDefCurrent.Prefix & DDTableCurrent.TableName, sfieldOLD), " WHERE [option] = '{F4EE0F9A-BA29-444D-B4FFD68A81730295}' "
                End If
            
            End If
        End If

        '<EhFooter>
        Exit Sub

SetDropdown_Err:
        MsgBox "The dynamic database is in error!" & Chr(13) & Chr(13) & "OASIS cannot find table: " & sField, vbCritical
        Exit Sub
        Resume Next
        'MsgBox "DynamicDataModule.SetDropdown_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Sub SetGridLookup(sField As String)
        '<EhHeader>
        On Error GoTo SetGridLookup_Err
        '</EhHeader>

        Dim oRS As ADODB.Recordset
        Dim sSQL As String
        Dim sLookupTableName As String

        If left(right$(sField, 6), 4) = "-PAD" Then
            sLookupTableName = left(sField, Len(sField) - 6)
        Else
            sLookupTableName = sField
        End If

        sSQL = "SELECT DISTINCT [GUID1], [option] FROM [" & sLookupTableName & "] order by [option]"
120     Set oRS = New ADODB.Recordset

122     oRS.Open sSQL, mConn, adOpenKeyset, adLockReadOnly ' ', adOpenStatic, adLockReadOnly

124     If Not dxDBGrid1.Columns.ColumnByFieldName(sField) Is Nothing Then

126         dxDBGrid1.Columns.ColumnByFieldName(sField).ColumnType = gedLookupEdit

128         With dxDBGrid1.Columns.ColumnByFieldName(sField).LookupColumn

130             .DisplaySize = mDDRSLinked.Fields.Item(sField).DefinedSize
132             .ListAutoWidth = True
134             .ListColumns = "*"
136             .ListFieldIndex = 0
138             .ListFieldName = "option"
140             .ListWidth = 0
142             .LookupCache = True
144             .LookupKeyField = "GUID1"
146             .LookupResultField = "option"
148             .ListFieldIndex = 0
150             .ListWidth = 200
152             .ListColumns = ""

            End With

154         Set dxDBGrid1.Columns.ColumnByFieldName(sField).LookupColumn.DataSource = oRS
156         dxDBGrid1.Columns.ColumnByFieldName(sField).caption = GetFieldCaption(mConn, mDDRSLinked, mDDRSLinked.Fields(sField).Name)

        End If

        '<EhFooter>
        Exit Sub

SetGridLookup_Err:
        MsgBox "SetGridLookup_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Function GetFieldCaption(oConn As ADODB.Connection, _
                                 RSLocalRecordset As ADODB.Recordset, _
                                 sFieldName As String)
        '<EhHeader>
        On Error GoTo GetFieldCaption_Err
        '</EhHeader>
 
        Dim oDB As ADOX.Catalog
        Dim itbl As ADOX.Table
        Dim fld As ADOX.Column
        
        Dim oRSCaption As ADODB.Recordset
        Dim bExitFor As Boolean
 
100     Set oDB = New ADOX.Catalog
102     Set itbl = New ADOX.Table
104     Set oDB.ActiveConnection = oConn
106     GetFieldCaption = "desc not defined"

108     For Each itbl In oDB.Tables

110         If itbl.Name = RSLocalRecordset.Fields(0).Properties(1) Then

112             If bSQLServerInUse Then
114                 Set oRSCaption = New ADODB.Recordset
116                 oRSCaption.Open "SELECT * FROM ::fn_listExtendedProperty " & "( 'MS_Description','user', 'dbo', 'table', '" & itbl.Name & "', 'column', '" & sFieldName & "')", mConn, adOpenDynamic, adLockBatchOptimistic

118                 If Not oRSCaption.EOF Then GetFieldCaption = oRSCaption.Fields("value").value
                    oRSCaption.Close
120                 Set oRSCaption = Nothing
122                 bExitFor = True
                Else
                
124                 GetFieldCaption = itbl.Columns(sFieldName).Properties(2).value
126                 bExitFor = True
                
                End If

128             If bExitFor Then Exit For
                
            End If

        Next
        
130     Set oDB = Nothing
132     Set itbl = Nothing
134     Set fld = Nothing

        '<EhFooter>
        Exit Function

GetFieldCaption_Err:
        MsgBox "DynamicDataModule.GetFieldCaption_Err (" & Erl & " ) " & Err.Description
        
        '</EhFooter>
End Function

Private Sub SetAccessRights()

    cmdAddData.Visible = DDTableCurrent.AllowAppend
    cmdEditData.Visible = True
    
    If DDTableCurrent.AllowEdit Then
        cmdEditData.caption = "Edit Data"
    Else
        cmdEditData.caption = "Browse Data"
    End If
    
End Sub

Private Function CreateCustomRS(oRS As ADODB.Recordset) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo CreateCustomRS_Err
        '</EhHeader>

        Dim sVisibleFields As String
        Dim sVisibleFieldsArray() As String
        Dim NewRS As New ADODB.Recordset
        Dim i As Long
    
100     If Not DDTableCurrent.IsDDTable Then
102         sVisibleFields = DDDefCurrent.Specification.Detail(DDDefCurrent.Specification.lCurrentRank - 1).sDataEntryFields
        Else
    
104         i = 0

106         Do Until i > UBound(DDDefCurrent.Specification.Detail)
    
108             If DDTableCurrent.TableName = DDDefCurrent.Specification.Detail(i).sTableName Then
110                 sVisibleFields = DDDefCurrent.Specification.Detail(i).sDataEntryFields
                    Exit Do
                End If

112             i = i + 1
            
            Loop
    
        End If

114     i = 0
        sVisibleFieldsArray = Split(sVisibleFields, ",")
        
116     'Do Until i = oRS.Fields.Count
        Do Until i > UBound(sVisibleFieldsArray)
            sVisibleFieldsArray(i) = Trim$(sVisibleFieldsArray(i))

118         'If InStr(UCase(sVisibleFields), UCase(oRS.Fields(i).Name)) > 0 Then
            If FieldExists(oRS, sVisibleFieldsArray(i)) Then

120             With oRS.Fields(sVisibleFieldsArray(i))
        
                    If .Type = adLongVarWChar Then
                        NewRS.Fields.Append .Name, adVarWChar, IIf(.DefinedSize > 8100, 8100, .DefinedSize) ', .Attributes
                    Else
                        NewRS.Fields.Append .Name, .Type, IIf(.DefinedSize > 8100, 8100, .DefinedSize), .Attributes
                    End If
    
122
                    
                End With
            
            End If

124         i = i + 1
        
        Loop
    
126     If Not NewRS.Fields.Count = 0 Then
    
128         NewRS.Open
130         NewRS.AddNew
132         CopyRSValues oRS, NewRS
134         Set CreateCustomRS = NewRS
    
        End If

        '<EhFooter>
        Exit Function

CreateCustomRS_Err:
        MsgBox "DynamicDataModule.CreateCustomRS_Err (" & Erl & " ) " & Err.Description
        
        '</EhFooter>
End Function

Private Sub ReturnToBrowseTab()

    'eCurrentOperation = Browsing
    UpdateContainerOfControls False
    C1TTab1Tab2.CurrTab = 0
    lblOperation.caption = "  Browsing records..."
    cmdCancel.Visible = False
    cmdNext.Visible = False
    cmdNext.caption = "Next"
    cmdSpatial.Visible = False
    cmdSpatial.Enabled = False
    ListDataElements.Enabled = True
    listDatabases.Enabled = True
    cmdAddData.Enabled = True
    cmdEditData.Enabled = True
    Call PointAtCurrentTableInList

End Sub

Public Sub ListDataElements_Click()
        '<EhHeader>
        On Error GoTo ListDataElements_Click_Err
        '</EhHeader>
        
        Dim sSQL As String
        Dim i As Long
        Dim sLockedArray() As String
        
100     C1TTab1Tab2.Visible = False

102     If Not IsNull(ListDataElements.Text) And Not ListDataElements.Text = "" Then
        
104         lblDescription.caption = "please wait momentarily for the data to load............"
106         lblDescription.FontSize = 15
108         'Set C1Elastic1.Picture = Nothing
            'C1Elastic1.BackColor = &H80&
110         'lblDescription.Picture = dxDBInspector1.BackgroundBitmap
112         ReturnToBrowseTab

114         DoEvents
116         bDisableEvents = True
        
            'Prepare mDDRS Recordset
118         sSQL = "SELECT * FROM [" & DDDefCurrent.Prefix & DDTableCurrent.TableName & "]"

120         If mDDRS.State = adStateOpen Then mDDRS.Close

122         Set mDDRS = New ADODB.Recordset
124         Set mDDRSManual = New ADODB.Recordset

126         If bSQLServerInUse Then mConn.CursorLocation = g_sGlobalCursorLocation
128         mDDRS.Open sSQL, mConn, adOpenDynamic, adLockBatchOptimistic
            
130         dxDBGrid1.Option = egoLoadAllRecords
132         dxDBGrid1.OptionEnabled = True
134         dxDBGrid1.Option = egoDynamicLoad
136         dxDBGrid1.OptionEnabled = False
138         dxDBGrid1.Option = egoAutoSort
140         dxDBGrid1.OptionEnabled = True
142         dxDBGrid1.Option = egoAutoSearchOnDynamicLoad
144         dxDBGrid1.OptionEnabled = True
146         Set mDDRSManual = CreateCustomRS(mDDRS)
        
            'Prepare Inspector
148         Set dxDBInspector1.DataSource = Nothing
150         C1Elastic1.Refresh
        
152         Set dxDBInspector1.DataSource = mDDRSManual
154         dxDBInspector1.ClearRows
156         dxDBInspector1.RetrieveFields

            'If Geotable prepare SQL Layer
158         If right$(DDDefCurrent.Prefix & DDTableCurrent.TableName, 4) = "_FEA" Then
160             Set mLayer = New TatukGIS_XDK10.XGIS_LayerSqlAdo
162             mLayer.Name = GUIDGen 'just give it a temp name
164             mLayer.SQLParameter("LAYER") = left$(DDDefCurrent.Prefix & DDTableCurrent.TableName, Len(DDDefCurrent.Prefix & DDTableCurrent.TableName) - 4)
166             mLayer.SQLParameter("DIALECT") = g_sGlobalDialect
168             mLayer.SQLParameter("ADO") = mConn.ConnectionString
170             mLayer.Open
172             SetGeoRsToCurrentUID
            End If
            
174         Call SetAccessRights
176         Call DisplayGridData

178         dxDBGrid1.Options.Unset (egoPreview)
180         dxDBGrid1.Options.Unset (egoAutoCalcPreviewLines)

182         If DDTableCurrent.IsMasterTable And DDDefCurrent.NumOfLinkedTables > 0 Then
184             'If Not mDDRS.EOF And Not mDDRS.Bof Then mDDRSGuid = mDDRS.Fields("GUID1").Value
186             UpdateLinkedTableCollection
            End If

188         listDatabases.Enabled = True
190         bDisableEvents = False
192         cmdCancel.caption = "Cancel"
        
194         If Not mDDRS.EOF Or Not mDDRS.Bof Then
        
196             mDDRS.MoveFirst
198             mDDRS.Find "[GUID1] = '" & dxDBGrid1.Columns(0).value & "'"
            
            End If
            
200         If DDTableCurrent.IsDDTable Then DDDefCurrent.Specification.lCurrentRank = 0
202         UpdateDescription

204         i = 0
206         PointAtCurrentTable DDTableCurrent.TableName
208         sLockedArray = Split(DDTableCurrentDetail.sLockedFields, ",")

210         Do Until i = UBound(sLockedArray) + 1 Or UBound(sLockedArray) = -1
            
212             If Not dxDBInspector1.RowByName(sLockedArray(i)) Is Nothing Then
214                 dxDBInspector1.RowByName(sLockedArray(i)).ReadOnly = True
216                 dxDBInspector1.RowByName(sLockedArray(i)).Alignment = taLeftJustify ' taRightJustify
                End If

218             i = i + 1
            Loop
            
220         Do Until i = UBound(DDNearbyFeatures) Or UBound(DDNearbyFeatures) = -1

222             If DDNearbyFeatures(i).sDataEntryTableName = DDTableCurrent.TableName Then
224                 If Not dxDBInspector1.RowByName(DDNearbyFeatures(i).sDataEntryFieldName) Is Nothing Then
                        
226                     dxDBInspector1.RowByName(DDNearbyFeatures(i).sDataEntryFieldName).ReadOnly = True
228                     dxDBInspector1.RowByName(DDNearbyFeatures(i).sDataEntryFieldName).Alignment = taLeftJustify ' taRightJustify
                        'dxDBInspector1.RowByName(DDNearbyFeatures(i).sDataEntryFieldName).color = vbGreen
                        dxDBInspector1.RowByName(DDNearbyFeatures(i).sDataEntryFieldName).ReadOnly = True
                        dxDBInspector1.RowByName(DDNearbyFeatures(i).sDataEntryFieldName).LookUpRow.ListRowCount = 1
                    End If
                End If

230             i = i + 1
            Loop

        End If

232     lblLabel1.Visible = True
234     dxProgressBar1.Visible = False
236     lblProgress.Visible = False
238     C1TTab1Tab2.Visible = True
240     cmdCancel.Visible = False
242     cmdNext.Visible = False

        '<EhFooter>
        Exit Sub

ListDataElements_Click_Err:
        MsgBox "This is an error in the design of the dynamic database! Error message: " & Err.Description, vbExclamation
        
        '</EhFooter>
End Sub

Private Function DoesQueryExist(oConn As ADODB.Connection, _
                                sQueryName As String) As Boolean

    Dim oRSQueries As New ADODB.Recordset

    With oRSQueries

        .Open oConn.OpenSchema(adSchemaViews)
        .Filter = "TABLE_NAME LIKE '" & sQueryName & "%'"

        If .EOF Then

            DoesQueryExist = False
        Else

            DoesQueryExist = True
        End If

    End With

    Set oRSQueries = Nothing

End Function

Private Sub UpdateDescription()
    Dim i As Long
 
    If Not DDTableCurrent.IsDDTable Then
        
        lblDescription.caption = DDDefCurrent.Specification.Detail(DDDefCurrent.Specification.lCurrentRank - 1).sDescription
        lblDescription.FontSize = DDDefCurrent.Specification.Detail(DDDefCurrent.Specification.lCurrentRank - 1).lDescFontSize
        
    Else
    
        i = 0

        Do Until i > UBound(DDDefCurrent.Specification.Detail)
    
            If DDTableCurrent.TableName = DDDefCurrent.Specification.Detail(i).sTableName Then

                lblDescription.caption = DDDefCurrent.Specification.Detail(i).sDescription
                lblDescription.FontSize = DDDefCurrent.Specification.Detail(i).lDescFontSize
                
            End If
    
            i = i + 1
 
        Loop
    
    End If
    
End Sub

Private Sub UpdateLinkedTableCollection()

    Dim l As Long
    Dim lCurrentLinkedTable As Long
    l = 0
    lCurrentLinkedTable = 0
    
    ReDim mDDRSLinkedCollection(DDDefCurrent.NumOfLinkedTables)
    
    Do Until l = UBound(DDDefCurrent.Tables) + 1
        
        If DDDefCurrent.Tables(l).IsLinkedTable Then

            mDDRSLinkedCollection(lCurrentLinkedTable).sTableName = DDDefCurrent.Tables(l).TableName
            Set mDDRSLinkedCollection(lCurrentLinkedTable).RS = New ADODB.Recordset
            lCurrentLinkedTable = lCurrentLinkedTable + 1
            
        End If
    
        l = l + 1
    Loop

End Sub

Private Function GetRSFromLinkedTablesCollection(sTableName As String) As ADODB.Recordset

    Dim l As Long
    l = 0
    
    Set GetRSFromLinkedTablesCollection = Nothing
    
    If DDDefCurrent.NumOfLinkedTables > 0 Then
        
        Do Until (l = UBound(mDDRSLinkedCollection)) Or (Not GetRSFromLinkedTablesCollection Is Nothing)
    
            If mDDRSLinkedCollection(l).sTableName = sTableName Then
                Set GetRSFromLinkedTablesCollection = mDDRSLinkedCollection(l).RS
            End If
        
            l = l + 1
        Loop
    
    End If

End Function

Private Sub SetRSInLinkedTablesCollection(sTableName As String, _
                                          oRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo SetRSInLinkedTablesCollection_Err
        '</EhHeader>

        Dim l As Long
100     l = 0
    
        If DDDefCurrent.NumOfLinkedTables > 0 Then
    
102         Do Until l = UBound(mDDRSLinkedCollection)
    
104             If mDDRSLinkedCollection(l).sTableName = sTableName Then
106                 Set mDDRSLinkedCollection(l).RS = oRS '.Clone
                End If
        
108             l = l + 1
            Loop
        
        End If

        '<EhFooter>
        Exit Sub

SetRSInLinkedTablesCollection_Err:
        MsgBox "DynamicDataModule.SetRSInLinkedTablesCollection_Err (" & Erl & " ) " & Err.Description
        'Stop
        '</EhFooter>
End Sub

Public Sub SaveChanges(bSave As Boolean, _
                       oLayer As TatukGIS_XDK10.XGIS_LayerVector, _
                       sKML As String)

    If bSave Then
        Set mShapeNew = oLayer.GetShape(oLayer.GetLastUid)
        sKMLString = sKML
        UpdateCreatedPoint mShapeNew.Centroid.X, mShapeNew.Centroid.Y
    Else
        Set mLayerCopy = Nothing
        sKMLString = ""
    End If

End Sub

Public Sub SaveNonSpatialXY(X As Double, _
                            Y As Double, _
                            sMGRS As String)
    
    If Not dxDBInspector1.RowByName("Longitude") Is Nothing Then dxDBInspector1.RowByName("Longitude").value = X
    
    If Not dxDBInspector1.RowByName("Latitude") Is Nothing Then dxDBInspector1.RowByName("Latitude").value = Y
    
    If Not dxDBInspector1.RowByName("MGRS") Is Nothing Then dxDBInspector1.RowByName("MGRS").value = sMGRS
    
    UpdateCreatedPoint X, Y

End Sub

Public Sub UpdateCreatedPoint(X As Double, _
                              Y As Double)
        '<EhHeader>
        On Error GoTo UpdateCreatedPoint_Err
        '</EhHeader>
    
        Dim i As Long
        Dim sLayerName As String
        Dim sLayerFieldName As String
        Dim sDataEntryFieldName As String
        Dim oLayer As XGIS_LayerAbstract
        Dim sResult As String
        Dim sFieldValue As String
        Dim dDistance As Double
    
100     If X = 0 And Y = 0 Then
102         Set mCreatedPoint = Nothing
        Else
104         Set mCreatedPoint = New XGIS_Point
106         mCreatedPoint.Prepare X, Y
        
108         i = 0

110         Do Until i = UBound(DDNearbyFeatures)
        
112             If DDTableCurrent.TableName = DDNearbyFeatures(i).sDataEntryTableName Then
            
114                 sLayerName = DDNearbyFeatures(i).sLayerName
116                 sLayerFieldName = DDNearbyFeatures(i).sLayerFieldName
118                 sDataEntryFieldName = DDNearbyFeatures(i).sDataEntryFieldName
120                 RaiseEvent GetNearestShapeInfo(mCreatedPoint, sLayerName, sLayerFieldName, sFieldValue, dDistance)
122                 dxDBInspector1.RowByName(sDataEntryFieldName).value = sFieldValue
124                 dxDBInspector1.RowByName(sDataEntryFieldName).ReadOnly = True
126                 dxDBInspector1.RowByName(sDataEntryFieldName).Alignment = taLeftJustify
                
130                 If DDNearbyFeatures(i).bDistance Then
132                     dxDBInspector1.RowByName(sDataEntryFieldName).value = sFieldValue & " [distance to: " & Round(dDistance, 2) & " km]"
                    End If
                
                End If
        
134             i = i + 1
            Loop
        
        End If

        '<EhFooter>
        Exit Sub

UpdateCreatedPoint_Err:
        'Err.Raise vbObjectError + 100, _
         "OASISClient.DynamicDataModule.UpdateCreatedPoint", _
         "DynamicDataModule component failure"
        '</EhFooter>
End Sub

Private Sub UpdateValidationMask(ByVal sTableName As String)
        '<EhHeader>
        On Error GoTo UpdateValidationMask_Err
        '</EhHeader>

        Dim i As Long
102     i = 0

104     Do Until i = UBound(DDValidation) 'Or bBreakout
        
106         If (sTableName = DDValidation(i).sTableName) And Not (dxDBInspector1.RowByName(DDValidation(i).sFieldName) Is Nothing) Then
            
108             With dxDBInspector1.RowByName(DDValidation(i).sFieldName)
                
110                 If Len(DDValidation(i).sEditMask) > 0 Then
112                     .RowType = iedMaskEdit
114                     .MaskRow.EditMask = DDValidation(i).sEditMask
                        
                    End If
        
                End With
                
            End If
        
116         i = i + 1
        Loop

        '<EhFooter>
        Exit Sub

UpdateValidationMask_Err:
        MsgBox "DynamicDataModule.UpdateValidationMask_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub

Private Function CheckValidation(ByVal sTableName As String) As Boolean
        '<EhHeader>
        On Error GoTo CheckValidation_Err
        '</EhHeader>
    
        Dim sMessage As String
        Dim sValidation() As String
        Dim i As Long
        Dim j As Long
        Dim numval As Double
100     i = 0
102     j = 0
    
104     CheckValidation = True

106     Do Until i = UBound(DDValidation) 'Or bBreakout
        
108         If (sTableName = DDValidation(i).sTableName) And Not (dxDBInspector1.RowByName(DDValidation(i).sFieldName) Is Nothing) Then
            
110             With dxDBInspector1.RowByName(DDValidation(i).sFieldName)
                
112                 If DDValidation(i).bRequired Then
                
114                     If Len(.value) = 0 Then
116                         CheckValidation = False

118                         If sMessage = "" Then sMessage = "The following validation errors were found in this form:" & vbCrLf
120                         sMessage = sMessage & vbCrLf & "field [" & .caption & "] should not be empty - please fill in this field!"
                        End If
                
                    End If
                
122                 If Len(DDValidation(i).sValidation) > 1 Then
                
124                     sValidation = Split(DDValidation(i).sValidation, "&&")
                    
126                     Do Until j > UBound(sValidation)
                   If Not IsNull(.value) And Not IsEmpty(.value) Then
128                         If Mid(sValidation(j), 2, 1) = "=" Then

130                             numval = CDbl(right(sValidation(j), Len(sValidation(j)) - 2))

132                             Select Case left(sValidation(j), 2)
                        
                                    Case ">="

134                                     If Not .value >= numval Then
136                                         If sMessage = "" Then sMessage = "The following validation errors were found in this form:" & vbCrLf
138                                         sMessage = sMessage & vbCrLf & "field [" & .caption & "] should be >= " & numval
140                                         CheckValidation = False

                                        End If
                                
142                                 Case "<="
                                
144                                     If Not .value <= numval Then
146                                         If sMessage = "" Then sMessage = "The following validation errors were found in this form:" & vbCrLf
148                                         sMessage = sMessage & vbCrLf & "field [" & .caption & "] should be <= " & numval
150                                         CheckValidation = False

                                        End If
                        
                                End Select
                        
                            Else
                        
152                             numval = CDbl(right(sValidation(j), Len(sValidation(j)) - 1))
                        
154                             Select Case left(sValidation(j), 1)
                        
                                    Case "<"
                                          
156                                     If Not .value < numval Then
158                                         If sMessage = "" Then sMessage = "The following validation errors were found in this form:" & vbCrLf
160                                         sMessage = sMessage & vbCrLf & "field [" & .caption & "] should be < " & numval
162                                         CheckValidation = False

                                        End If

164                                 Case ">"
                                  
166                                     If Not .value > numval Then
168                                         If sMessage = "" Then sMessage = "The following validation errors were found in this form:" & vbCrLf
170                                         sMessage = sMessage & vbCrLf & "field [" & .caption & "] should be > " & numval
172                                         CheckValidation = False

                                        End If

                                End Select
                        
                            End If
                            End If
                        
174                         j = j + 1
                        Loop
                
                    End If
                
                    '  = DDValidation(i).bRequired
                    ' sValidation = DDValidation(i).sValidation

                End With
                
            End If
        
176         i = i + 1
        Loop
    
178     If CheckValidation = False Then MsgBox sMessage, vbInformation, "Validation error!"

        '<EhFooter>
        Exit Function

CheckValidation_Err:
        CheckValidation = False
        MsgBox "DynamicDataModule.CheckValidation_Err (" & Erl & " ) " & Err.Description
        MsgBox sMessage, vbInformation, "Validation error!"
        '</EhFooter>
End Function

Private Function CheckValidationByNode(ByVal OldNode As DXDBINSPLibCtl.IdxDBRowNode, _
                                       ByVal sTableName As String) As Boolean
    'dude
    
    Dim sMessage As String
    Dim sValidation() As String
    Dim i As Long
    Dim j As Long
    Dim numval As Double
    i = 0
    j = 0
    
    CheckValidationByNode = True

    Do Until i = UBound(DDValidation) 'Or bBreakout
        
        If (sTableName = DDValidation(i).sTableName) And Not (dxDBInspector1.RowByName(DDValidation(i).sFieldName) Is Nothing) Then
            
            If DDValidation(i).sFieldName = OldNode.Row.FieldName Then
            
                With dxDBInspector1.RowByName(DDValidation(i).sFieldName)
                
                    'If DDValidation(i).bRequired Then
                
                    ' If Len(.value) = 0 Then
                    ' CheckValidationByNode = False

                    ' If sMessage = "" Then sMessage = "The following validation errors were found in this form:" & vbCrLf
                    ' sMessage = sMessage & vbCrLf & "field [" & .caption & "] should not be empty - please fill in this field!"
                    '  End If
                
                    '  End If
                
                    If Len(DDValidation(i).sValidation) > 1 Then
                
                        sValidation = Split(DDValidation(i).sValidation, "&&")
                    
                        Do Until j > UBound(sValidation)
                    If Not IsNull(.value) And Not IsEmpty(.value) Then
                            If Mid(sValidation(j), 2, 1) = "=" Then

                                numval = CDbl(right(sValidation(j), Len(sValidation(j)) - 2))

                                Select Case left(sValidation(j), 2)
                        
                                    Case ">="

                                        If Not .value >= numval Then
                                            If sMessage = "" Then sMessage = "The following validation errors were found in this form:" & vbCrLf
                                            sMessage = sMessage & vbCrLf & "field [" & .caption & "] should be >= " & numval
                                            CheckValidationByNode = False

                                        End If
                                
                                    Case "<="
                                
                                        If Not .value <= numval Then
                                            If sMessage = "" Then sMessage = "The following validation errors were found in this form:" & vbCrLf
                                            sMessage = sMessage & vbCrLf & "field [" & .caption & "] should be <= " & numval
                                            CheckValidationByNode = False

                                        End If
                        
                                End Select
                        
                            Else
                        
                                numval = CDbl(right(sValidation(j), Len(sValidation(j)) - 1))
                        'MsgBox mDDRS.Fields(.FieldName).value
                                Select Case left(sValidation(j), 1)
                        
                                    Case "<"
                                          
                                        If Not .value < numval Then
                                            If sMessage = "" Then sMessage = "The following validation errors were found in this form:" & vbCrLf
                                            sMessage = sMessage & vbCrLf & "field [" & .caption & "] should be < " & numval
                                            CheckValidationByNode = False

                                        End If

                                    Case ">"
                                  
                                        If Not .value > numval Then
                                            If sMessage = "" Then sMessage = "The following validation errors were found in this form:" & vbCrLf
                                            sMessage = sMessage & vbCrLf & "field [" & .caption & "] should be > " & numval
                                            CheckValidationByNode = False

                                        End If

                                End Select
                        
                            End If
                        End If
                            j = j + 1
                        Loop
                
                    End If
                
                    '  = DDValidation(i).bRequired
                    ' sValidation = DDValidation(i).sValidation

                End With

            End If
        End If
        
        i = i + 1
    Loop
    
    If CheckValidationByNode = False Then MsgBox sMessage, vbInformation, "Validation error!"

End Function

Private Sub GetDDDetailedSpec(sTablePrefix As String, _
                              oDynamicDataDef As DYNAMIC_DATA_DEF, _
                              sConnectionString As String)
        '<EhHeader>
        On Error GoTo GetDDDetailedSpec_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
        Dim oCn As New ADODB.Connection
        Dim i As Long
        Dim j As Long
        Dim sExcludedFields As String
        Dim sExcludedArray() As String
        Dim sLockedFields As String
        Dim sLockedArray() As String
        Dim sLockedField As String
        
        oDynamicDataDef.NumMainDataEntryPages = 0
100     i = 0

102     If bSQLServerInUse Then
104         oCn.Open g_sGlobalConnectionString
        Else
106         oCn.Open sConnectionString
        End If
    
108     With oRS
    
110         .Open "SELECT * FROM [" & sTablePrefix & "Specification] ORDER BY [lRank]", oCn, adOpenDynamic, adLockBatchOptimistic
        
112         If Not .EOF Or Not .Bof Then
        
114             .MoveFirst

116             Do Until .EOF

118                 ReDim Preserve oDynamicDataDef.Specification.Detail(i + 1)
120                 oDynamicDataDef.Specification.Detail(i).bIsLinkedTable = .Fields("bIsLinkedTable").value
122                 oDynamicDataDef.Specification.Detail(i).bIsMasterTable = .Fields("bIsMasterTable").value
124                 oDynamicDataDef.Specification.Detail(i).lRank = .Fields("lRank").value

                    If oDynamicDataDef.Specification.Detail(i).bIsLinkedTable Or oDynamicDataDef.Specification.Detail(i).bIsMasterTable Then
                        oDynamicDataDef.NumMainDataEntryPages = oDynamicDataDef.NumMainDataEntryPages + 1
                    End If
    
126                 If .Fields("bIsLinkedTable").value Or .Fields("bIsMasterTable").value Then oDynamicDataDef.Specification.lMaxRank = .Fields("lRank").value
128                 If Len(.Fields("sCaption").value) > 1 Then oDynamicDataDef.Specification.Detail(i).sCaption = .Fields("sCaption").value

130                 oDynamicDataDef.Specification.Detail(i).sDescription = .Fields("sDescription").value
132                 oDynamicDataDef.Specification.Detail(i).lDescFontSize = .Fields("lDescFontSize").value
134                 oDynamicDataDef.Specification.Detail(i).sTableName = .Fields("sTableName").value
136                 oDynamicDataDef.Specification.Detail(i).sSQLView = " " & .Fields("sGridQuery").value
138                 oDynamicDataDef.Specification.Detail(i).sSQLViewMSSQL = " " & .Fields("sGridQueryMSSQL").value
                    
140                 sExcludedFields = oDynamicDataDef.ExcludedFields
142                 sExcludedArray = Split(sExcludedFields, ";")
                    
144                 j = 0

146                 Do Until j = UBound(sExcludedArray) + 1 Or UBound(sExcludedArray) = -1
                    
148                     If InStr(left(sExcludedArray(j), Len(.Fields("sTableName").value)), .Fields("sTableName").value) > 0 Then
150                         .Fields("sDataEntryFields").value = Replace(.Fields("sDataEntryFields").value, right(sExcludedArray(j), Len(sExcludedArray(j)) - 1 - Len(.Fields("sTableName").value)), "")

152                         If left$(.Fields("sDataEntryFields").value, 1) = "," Then .Fields("sDataEntryFields").value = right$(.Fields("sDataEntryFields").value, Len(.Fields("sDataEntryFields").value) - 1)
                        End If

154                     j = j + 1
                    Loop
                    
156                 sLockedFields = oDynamicDataDef.LockedFields
158                 sLockedArray = Split(sLockedFields, ";")
160                 j = 0
162                 oDynamicDataDef.Specification.Detail(i).sLockedFields = ""

164                 Do Until j = UBound(sLockedArray) + 1 Or UBound(sLockedArray) = -1
                    
166                     If InStr(left(sLockedArray(j), Len(.Fields("sTableName").value)), .Fields("sTableName").value) > 0 Then
168                         sLockedField = right(sLockedArray(j), Len(sLockedArray(j)) - 1 - Len(.Fields("sTableName").value))
170                         oDynamicDataDef.Specification.Detail(i).sLockedFields = oDynamicDataDef.Specification.Detail(i).sLockedFields & sLockedField & ","
                        End If

172                     j = j + 1
                    Loop
                    
174                 oDynamicDataDef.Specification.Detail(i).sDataEntryFields = .Fields("sDataEntryFields").value
176                 i = i + 1
178                 .MoveNext
                Loop
            
180             oDynamicDataDef.Specification.lCurrentRank = 0
182             .Close
        
            End If
    
        End With
    
184     Set oRS = Nothing
186     Set oCn = Nothing
                              
        '<EhFooter>
        Exit Sub

GetDDDetailedSpec_Err:
        Set oRS = Nothing
        Set oCn = Nothing
        MsgBox "DynamicDataModule.GetDDDetailedSpec_Err (" & Erl & " ) " & Err.Description
        '</EhFooter>
End Sub
 
Private Sub GetDDDetailedInfo(sTablePrefix As String, _
                              sAccessRights As String, _
                              DynamicDataDef As DYNAMIC_DATA_DEF)
        '<EhHeader>
        On Error GoTo GetDDDetailedInfo_Err
        '</EhHeader>

        Dim i As Long
        Dim l As Long
        Dim sTableName As String
        Dim sAccessRightsArray() As String
        Dim sSmallSplit() As String
    
100     ReDim Preserve DynamicDataDef.Tables(0)
102     DynamicDataDef.NumOfLinkedTables = 0
103     DynamicDataDef.NumOfDDTables = 0
104     DynamicDataDef.LinkedTableNames = ""
106     sAccessRightsArray = Split(sAccessRights, ";", -1, vbTextCompare)

112     i = 0

114     Do Until i = UBound(sAccessRightsArray) Or UBound(sAccessRightsArray) = -1
    
116         ReDim Preserve DynamicDataDef.Tables(UBound(DynamicDataDef.Tables) + 1)
        
118         With DynamicDataDef.Tables(UBound(DynamicDataDef.Tables))
        
120             sSmallSplit = Split(sAccessRightsArray(i), ",", -1, vbTextCompare)
            
122             .TableName = sSmallSplit(0)
124             .AllowAppend = False
126             .AllowDelete = False
128             .AllowEdit = False
130             .AllowRead = False

132             If left$(.TableName, 4) = "link" Then
134                 .IsLinkedTable = True
136                 DynamicDataDef.LinkedTableNames = DynamicDataDef.LinkedTableNames & "," & .TableName
138                 DynamicDataDef.NumOfLinkedTables = DynamicDataDef.NumOfLinkedTables + 1
140             ElseIf .TableName = "mastertable" Then
142                 .IsMasterTable = True
                Else
144                 .IsDDTable = True
                    DynamicDataDef.NumOfDDTables = DynamicDataDef.NumOfDDTables + 1
146                 DynamicDataDef.DDTableNames = DynamicDataDef.DDTableNames & "," & .TableName
                End If
                
                .IsGEOTable = IIf(right$(.TableName, 4) = "_FEA", True, False)

148             If InStr(1, sSmallSplit(1), "r", vbTextCompare) Then .AllowRead = True
150             If InStr(1, sSmallSplit(1), "e", vbTextCompare) Then .AllowEdit = True
152             If InStr(1, sSmallSplit(1), "a", vbTextCompare) Then .AllowAppend = True
154             If InStr(1, sSmallSplit(1), "d", vbTextCompare) Then .AllowDelete = True
            
            End With
            
196         i = i + 1
        Loop
        
198     If DynamicDataDef.NumOfLinkedTables > 0 Then
200         DynamicDataDef.LinkedTableNames = right$(DynamicDataDef.LinkedTableNames, Len(DynamicDataDef.LinkedTableNames) - 1)
        End If
    
202     If Len(DynamicDataDef.DDTableNames) > 0 Then
204         DynamicDataDef.DDTableNames = right$(DynamicDataDef.DDTableNames, Len(DynamicDataDef.DDTableNames) - 1)
        End If

        '<EhFooter>
        Exit Sub

GetDDDetailedInfo_Err:
        MsgBox "DynamicDataModule.GetDDDetailedInfo_Err (" & Erl & " ) " & Err.Description
        Resume Next
        '</EhFooter>
End Sub

Private Function GetDDTableNamesWithField(sFieldName As String) As String
        '<EhHeader>
        On Error GoTo GetDDTableNamesWithField_Err
        '</EhHeader>

        Dim sDDTableNamesLocal() As String
        Dim i As Long
    
100     sDDTableNamesLocal = Split(DDDefCurrent.DDTableNames, ",")
102     sLinkedTableNames = Split(DDDefCurrent.LinkedTableNames, ",")
    
104     GetDDTableNamesWithField = ""
    
106     i = 0

108     Do Until i > UBound(sDDTableNamesLocal)
    
110         If CheckIfFieldInTable(sFieldName, DDDefCurrent.Prefix & sDDTableNamesLocal(i)) Then
112             GetDDTableNamesWithField = GetDDTableNamesWithField & "," & sDDTableNamesLocal(i)
            End If

114         i = i + 1
        Loop
    
116     i = 0

118     Do Until i > UBound(sLinkedTableNames)
    
120         If CheckIfFieldInTable(sFieldName, DDDefCurrent.Prefix & sLinkedTableNames(i)) Then
122             GetDDTableNamesWithField = GetDDTableNamesWithField & "," & sLinkedTableNames(i)
            End If

124         i = i + 1
        Loop
    
126     If CheckIfFieldInTable(sFieldName, DDDefCurrent.Prefix & "mastertable") Then
128         GetDDTableNamesWithField = GetDDTableNamesWithField & ",mastertable"
        End If
    
130     If Len(GetDDTableNamesWithField) > 0 Then
132         GetDDTableNamesWithField = right$(GetDDTableNamesWithField, Len(GetDDTableNamesWithField) - 1)
        End If
    
        '<EhFooter>
        Exit Function

GetDDTableNamesWithField_Err:
        MsgBox "DynamicDataModule.GetDDTableNamesWithField_Err (" & Erl & " ) " & Err.Description
        'Resume Next
        '</EhFooter>
End Function

Function CheckIfFieldInTable(sFieldName As String, sTableName As String) As Boolean
        '<EhHeader>
        On Error GoTo CheckIfFieldInTable_Err
        '</EhHeader>
    
        Dim cat As New ADOX.Catalog 'Root object of ADOX.
        Dim tbl As ADOX.Table       'Each Table in Tables.
        Dim Col As ADOX.Column      'Each Column in the Table.
        Dim prp As ADOX.Property

100     Set cat.ActiveConnection = mConn
102     Set tbl = cat.Tables(sTableName)
    
104     CheckIfFieldInTable = False

106     For Each Col In tbl.Columns
        
108         If Not CheckIfFieldInTable Then
        
110             If Col.Name = sFieldName Then CheckIfFieldInTable = True

            End If
        
        Next

112     Set prp = Nothing
114     Set Col = Nothing
116     Set tbl = Nothing
118     Set cat = Nothing
    
        '<EhFooter>
        Exit Function

CheckIfFieldInTable_Err:
        CheckIfFieldInTable = False
        '</EhFooter>
End Function

Private Sub dxDBInspector1_OnExpanding(ByVal RowNode As DXDBINSPLibCtl.IdxDBRowNode, Allow As Boolean)
Stop
End Sub

Private Sub dxDBInspector1_OnKeyDown(KeyCode As Integer, _
                                     ByVal Shift As Integer)
                                     
    If KeyCode = vbKeyEscape Then KeyCode = 0
 
End Sub

Private Sub dxDBInspector1_OnKeyUp(KeyCode As Integer, _
                                   ByVal Shift As Integer)
    'SimulateRightArrow
    If dxDBInspector1.Count > 0 Then
    If dxDBInspector1.FocusedNode.Row.ReadOnly Or KeyCode = vbKeyEscape Then
    
        KeyCode = 0
        
    ElseIf KeyCode = vbKeyDelete And dxDBInspector1.FocusedNode.Row.RowType = 11 Then
    
        dxDBInspector1.FocusedNode.Row.value = ""
        dxDBInspector1.EndUpdate
        dxDBInspector1.BeginUpdate
        KeyCode = 0
        
        ' ElseIf KeyCode = vbKeyDelete Then
        ' KeyCode = 8
        
    ElseIf dxDBInspector1.FocusedNode.Row.RowType = iedMaskEdit Then
        
        If Len(dxDBInspector1.FocusedNode.Row.MaskRow.EditMask) > 0 Then
        
            If KeyCode = vbKeyBack Then
                On Error Resume Next
                dxDBInspector1.PostEditor
                DebugPrint "edit: " & dxDBInspector1.FocusedNode.Row.EditText
                DebugPrint "disp: " & dxDBInspector1.FocusedNode.Row.DisplayText
                DebugPrint "valu: " & dxDBInspector1.FocusedNode.Row.value
                DebugPrint "EditMask: " & dxDBInspector1.FocusedNode.Row.MaskRow.EditMask
                
                dxDBInspector1.FocusedNode.Row.EditText = ""
                dxDBInspector1.FocusedNode.Row.value = ""
                KeyCode = 0
            End If
        
        End If

    Else
    
        'On Error Resume Next
        'DebugPrint "KeyAscii: " & KeyAscii
        'If dxDBInspector1.FocusedNode.Row.RowType = iedMaskEdit And KeyAscii = vbKeyBack Then KeyAscii = vbKeyEscape

        If dxDBInspector1.FocusedNode.Row.ReadOnly Then
            KeyCode = 0
    
        End If
    
        If KeyCode = vbKeyEscape Then
    
            dxDBInspector1.FocusedNode.Row.EditText = ""
            KeyCode = 0
        
        Else
   
            If KeyCode = vbKeyBack Then
                dxDBInspector1.PostEditor
            End If

            If KeyCode = vbKeyBack And Len(dxDBInspector1.FocusedNode.Row.value) > 1 And Not dxDBInspector1.FocusedNode.Row.RowType = iedDateEdit Then
        
                dxDBInspector1.FocusedNode.Row.value = left(dxDBInspector1.FocusedNode.Row.value, Len(dxDBInspector1.FocusedNode.Row.value) - 1)
                dxDBInspector1.EndUpdate
                dxDBInspector1.BeginUpdate
            
                KeyCode = 0
            ElseIf KeyCode = vbKeyBack Then

                If IsNumeric(dxDBInspector1.FocusedNode.Row.value) Then
                    dxDBInspector1.FocusedNode.Row.value = Null
                ElseIf dxDBInspector1.FocusedNode.Row.RowType = iedDateEdit Then
                    dxDBInspector1.FocusedNode.Row.DateRow.DateOnError = deNull
                    dxDBInspector1.FocusedNode.Row.value = deNull
                    MsgBox "Please note that the current value in the date field will be saved to the database as a NULL value.  If you were meant to clear this field this message is to confirm that it has been done and please ignore the date now in this field", vbInformation
                Else
                    dxDBInspector1.FocusedNode.Row.value = ""
                End If
            
                'dxDBInspector1.FocusedNode.Expand True
            
                KeyCode = 0
            End If
    
        End If
    
    End If
    End If
    
End Sub

Private Sub ExportToHTML_Click()

    Dim c As New cCommonDialog

    With c
        .CancelError = False
        .DialogTitle = "Export data to..."
        .Filter = "HTML (.html)|*.html"
        .DefaultExt = ".html"
        .ShowSave
    End With
            
    If c.ConfirmOK And Not c.Filename = "" And Not IsNull(c.Filename) Then
        dxDBGrid1.M.ExportToHTML c.Filename
    End If

End Sub

Private Sub ExportToXLS_Click()

    Dim c As New cCommonDialog

    With c
        .CancelError = False
        .DialogTitle = "Export data to..."
        .Filter = "Microsoft Excel (.xls)|*.xls"
        .DefaultExt = ".xls"
        .ShowSave
    End With
            
    If c.ConfirmOK And Not c.Filename = "" And Not IsNull(c.Filename) Then
        dxDBGrid1.M.ExportToXLS c.Filename
    End If
    
End Sub

Private Sub ExportToXML_Click()

    Dim c As New cCommonDialog

    With c
        .CancelError = False
        .DialogTitle = "Export data to..."
        .Filter = "XML (.xml)|*.xml"
        .DefaultExt = ".xml"
        .ShowSave
    End With
            
    If c.ConfirmOK And Not c.Filename = "" And Not IsNull(c.Filename) Then
        dxDBGrid1.M.ExportToXML c.Filename
    End If
    
End Sub

Private Sub UserControl_Terminate()

    If cmdCancel.Visible Then Call cmdCancel_Click
    
End Sub



