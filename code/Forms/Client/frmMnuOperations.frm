VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{9C989D2F-3596-477B-B719-6DC4E6893AF2}#1.0#0"; "TatukGIS_XDK10_WIN32.ocx"
Begin VB.Form frmMnuOperations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operations Properties"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin TatukGIS_XDK10.XGIS_ViewerWnd GISSecurity 
      Height          =   1515
      Left            =   1140
      TabIndex        =   94
      Top             =   1500
      Visible         =   0   'False
      Width           =   1515
      BigExtentMargin =   -10
      RestrictedDrag  =   -1  'True
      CachedPaint     =   -1  'True
      IncrementalPaint=   -1  'True
      FullPaint       =   -1  'True
      CodePage        =   0
      OutCodePage     =   0
      CharSet         =   0
      UseRTree        =   0   'False
      PrinterTileSize =   2700
      PrintTitle      =   ""
      PrintSubtitle   =   ""
      PrintFooter     =   ""
      BeginProperty PrintTitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PrintSubtitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PrintFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PrintTitleFontColor=   -16777208
      PrintSubtitleFontColor=   -16777208
      PrintFooterFontColor=   -16777208
      SelectionColor  =   16777215
      SelectionPattern=   "frmMnuOperations.frx":0000
      SelectionTransparency=   100
      SelectionWidth  =   100
      SelectionOutlineOnly=   0   'False
      OldCachedPaint  =   0   'False
      PrinterModeDraft=   0   'False
      PrinterModeForceBitmap=   0   'False
      GDIType         =   0
      ScaleAsFloat    =   1
      Mode            =   0
      BorderStyle     =   1
      CursorForDrag   =   0
      CursorForSelect =   0
      CursorForZoom   =   0
      CursorForEdit   =   0
      MinZoomSize     =   -5
      ScrollBars      =   0
      AutoCenter      =   0   'False
      Align           =   0
      Ctl3D           =   -1  'True
      Object.Visible         =   -1  'True
      Cursor          =   16
      DoubleBuffered  =   0   'False
      ModeMouseButton =   0
      CursorForUserDefined=   0
      View3D          =   0   'False
   End
   Begin TatukGIS_XDK10.XGIS_ViewerWnd GIS 
      Height          =   1515
      Left            =   1140
      TabIndex        =   93
      Top             =   2880
      Visible         =   0   'False
      Width           =   1515
      BigExtentMargin =   -10
      RestrictedDrag  =   -1  'True
      CachedPaint     =   -1  'True
      IncrementalPaint=   -1  'True
      FullPaint       =   -1  'True
      CodePage        =   0
      OutCodePage     =   0
      CharSet         =   0
      UseRTree        =   0   'False
      PrinterTileSize =   2700
      PrintTitle      =   ""
      PrintSubtitle   =   ""
      PrintFooter     =   ""
      BeginProperty PrintTitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PrintSubtitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PrintFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PrintTitleFontColor=   -16777208
      PrintSubtitleFontColor=   -16777208
      PrintFooterFontColor=   -16777208
      SelectionColor  =   16777215
      SelectionPattern=   "frmMnuOperations.frx":0062
      SelectionTransparency=   100
      SelectionWidth  =   100
      SelectionOutlineOnly=   0   'False
      OldCachedPaint  =   0   'False
      PrinterModeDraft=   0   'False
      PrinterModeForceBitmap=   0   'False
      GDIType         =   0
      ScaleAsFloat    =   1
      Mode            =   0
      BorderStyle     =   1
      CursorForDrag   =   0
      CursorForSelect =   0
      CursorForZoom   =   0
      CursorForEdit   =   0
      MinZoomSize     =   -5
      ScrollBars      =   0
      AutoCenter      =   0   'False
      Align           =   0
      Ctl3D           =   -1  'True
      Object.Visible         =   -1  'True
      Cursor          =   16
      DoubleBuffered  =   0   'False
      ModeMouseButton =   0
      CursorForUserDefined=   0
      View3D          =   0   'False
   End
   Begin VB.CommandButton cmdExportLegend 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   90
      MaskColor       =   &H00000000&
      Picture         =   "frmMnuOperations.frx":00C4
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "Add Legend To Clipboard"
      Top             =   420
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   405
   End
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   5685
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3735
      _cx             =   6588
      _cy             =   10028
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
      Align           =   5
      AutoSizeChildren=   6
      BorderWidth     =   1
      ChildSpacing    =   1
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
      Begin C1SizerLibCtl.C1Tab C1TabOpsView 
         Height          =   5655
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   3705
         _cx             =   6535
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
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "Legend|Security|Mine Action|LIS|W3|Settings|Map Library|NFI|Search"
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
         Flags(4)        =   2
         Flags(5)        =   2
         Flags(7)        =   2
         Flags(8)        =   2
         Begin C1SizerLibCtl.C1Elastic elNFI 
            Height          =   5280
            Left            =   6150
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   330
            Width           =   3615
            _cx             =   6376
            _cy             =   9313
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
            GridRows        =   3
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmMnuOperations.frx":07C6
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic elAdmin 
               Height          =   660
               Left            =   30
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   30
               Width           =   3555
               _cx             =   6271
               _cy             =   1164
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
               Begin VB.ComboBox comGovernorates 
                  Height          =   315
                  Left            =   135
                  Style           =   2  'Dropdown List
                  TabIndex        =   60
                  Top             =   315
                  Width           =   3330
               End
               Begin VB.Label lblChooseLocation 
                  Caption         =   "Choose Location:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   62
                  Top             =   45
                  Width           =   1500
               End
            End
            Begin C1SizerLibCtl.C1Elastic elNFINav 
               Height          =   495
               Left            =   30
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   4755
               Width           =   3555
               _cx             =   6271
               _cy             =   873
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
               Begin VB.CommandButton cmdIOMJOC 
                  Caption         =   "IOM JOC"
                  Height          =   285
                  Left            =   360
                  TabIndex        =   68
                  Top             =   135
                  Width           =   1140
               End
            End
            Begin DXDBGRIDLibCtl.dxDBGrid dxTypeScoring 
               Height          =   4065
               Left            =   30
               OleObjectBlob   =   "frmMnuOperations.frx":0813
               TabIndex        =   61
               Top             =   690
               Width           =   3555
            End
         End
         Begin C1SizerLibCtl.C1Elastic elMapLibrary 
            Height          =   5280
            Left            =   5850
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   330
            Width           =   3615
            _cx             =   6376
            _cy             =   9313
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
            BorderWidth     =   6
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
            GridRows        =   1
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmMnuOperations.frx":24F3
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin OASISClient.ctlMapFileContainer ctlMapFileContainer1 
               Height          =   5100
               Left            =   90
               TabIndex        =   98
               Top             =   90
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   8996
            End
         End
         Begin C1SizerLibCtl.C1Elastic elMineAction 
            Height          =   5280
            Left            =   4650
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   330
            Width           =   3615
            _cx             =   6376
            _cy             =   9313
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
            BorderWidth     =   1
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
            _GridInfo       =   $"frmMnuOperations.frx":252B
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Tab tabMAOptions 
               Height          =   4770
               Left            =   15
               TabIndex        =   23
               Top             =   15
               Width           =   3585
               _cx             =   6324
               _cy             =   8414
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
               TabOutlineColor =   -2147483632
               FrontTabForeColor=   -2147483630
               Caption         =   "Tab&1|Tab&2|Tab&3"
               Align           =   0
               CurrTab         =   0
               FirstTab        =   0
               Style           =   3
               Position        =   0
               AutoSwitch      =   -1  'True
               AutoScroll      =   -1  'True
               TabPreview      =   -1  'True
               ShowFocusRect   =   0   'False
               TabsPerPage     =   0
               BorderWidth     =   0
               BoldCurrent     =   0   'False
               DogEars         =   -1  'True
               MultiRow        =   0   'False
               MultiRowOffset  =   200
               CaptionStyle    =   0
               TabHeight       =   1
               TabCaptionPos   =   5
               TabPicturePos   =   0
               CaptionEmpty    =   ""
               Separators      =   0   'False
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   37
               Begin C1SizerLibCtl.C1Elastic elTab3 
                  Height          =   4740
                  Left            =   4500
                  TabIndex        =   35
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   3555
                  _cx             =   6271
                  _cy             =   8361
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
                  GridCols        =   1
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"frmMnuOperations.frx":256D
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin C1SizerLibCtl.C1Elastic elMAHolder3 
                     Height          =   4740
                     Left            =   0
                     TabIndex        =   36
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   3555
                     _cx             =   6271
                     _cy             =   8361
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
                     Begin VB.Frame FraSettings 
                        Caption         =   "Settings:"
                        Height          =   4335
                        Left            =   90
                        TabIndex        =   37
                        Top             =   180
                        Width           =   4065
                     End
                  End
               End
               Begin C1SizerLibCtl.C1Elastic elTab2 
                  Height          =   4740
                  Left            =   4200
                  TabIndex        =   32
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   3555
                  _cx             =   6271
                  _cy             =   8361
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
                  _GridInfo       =   $"frmMnuOperations.frx":259F
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin C1SizerLibCtl.C1Elastic elMapHolder2 
                     Height          =   4740
                     Left            =   0
                     TabIndex        =   33
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   3555
                     _cx             =   6271
                     _cy             =   8361
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
                     GridRows        =   1
                     GridCols        =   1
                     Frame           =   3
                     FrameStyle      =   0
                     FrameWidth      =   1
                     FrameColor      =   -2147483628
                     FrameShadow     =   -2147483632
                     FloodStyle      =   1
                     _GridInfo       =   $"frmMnuOperations.frx":25DC
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin C1SizerLibCtl.C1Elastic elUpdateMAData 
                        Height          =   4680
                        Left            =   30
                        TabIndex        =   34
                        TabStop         =   0   'False
                        Top             =   30
                        Width           =   3495
                        _cx             =   6165
                        _cy             =   8255
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
                        GridRows        =   4
                        GridCols        =   3
                        Frame           =   3
                        FrameStyle      =   0
                        FrameWidth      =   1
                        FrameColor      =   -2147483628
                        FrameShadow     =   -2147483632
                        FloodStyle      =   1
                        _GridInfo       =   $"frmMnuOperations.frx":2612
                        AccessibleName  =   ""
                        AccessibleDescription=   ""
                        AccessibleValue =   ""
                        AccessibleRole  =   9
                        Begin VB.Frame FraUpdateRegions 
                           Caption         =   "Available Regions"
                           Height          =   1290
                           Left            =   0
                           TabIndex        =   45
                           Top             =   3285
                           Width           =   3495
                           Begin VB.OptionButton OptRegion1 
                              Caption         =   "All Regions"
                              Height          =   195
                              Index           =   5
                              Left            =   2115
                              TabIndex        =   46
                              Top             =   450
                              Width           =   1185
                           End
                           Begin VB.OptionButton OptRegion1 
                              Caption         =   "Region 5"
                              Height          =   195
                              Index           =   4
                              Left            =   2115
                              TabIndex        =   47
                              Top             =   225
                              Width           =   1185
                           End
                           Begin VB.OptionButton OptRegion1 
                              Caption         =   "Region 4"
                              Height          =   195
                              Index           =   3
                              Left            =   1125
                              TabIndex        =   48
                              Top             =   450
                              Width           =   1185
                           End
                           Begin VB.OptionButton OptRegion1 
                              Caption         =   "Region 3"
                              Height          =   195
                              Index           =   2
                              Left            =   1125
                              TabIndex        =   49
                              Top             =   225
                              Width           =   1185
                           End
                           Begin VB.OptionButton OptRegion1 
                              Caption         =   "Region 1"
                              Height          =   195
                              Index           =   0
                              Left            =   90
                              TabIndex        =   51
                              Top             =   225
                              Value           =   -1  'True
                              Width           =   1185
                           End
                           Begin VB.OptionButton OptRegion1 
                              Caption         =   "Region 2"
                              Height          =   195
                              Index           =   1
                              Left            =   90
                              TabIndex        =   50
                              Top             =   450
                              Width           =   1185
                           End
                        End
                        Begin MSComctlLib.ListView lstUpdates 
                           Height          =   4125
                           Left            =   0
                           TabIndex        =   52
                           Top             =   450
                           Width           =   3495
                           _ExtentX        =   6165
                           _ExtentY        =   7276
                           View            =   3
                           LabelEdit       =   1
                           Sorted          =   -1  'True
                           LabelWrap       =   -1  'True
                           HideSelection   =   -1  'True
                           Checkboxes      =   -1  'True
                           _Version        =   393217
                           ForeColor       =   -2147483640
                           BackColor       =   -2147483643
                           BorderStyle     =   1
                           Appearance      =   1
                           BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                              Name            =   "MS Serif"
                              Size            =   6.75
                              Charset         =   0
                              Weight          =   700
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           NumItems        =   3
                           BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                              Text            =   "IMSMA Data"
                              Object.Width           =   3528
                           EndProperty
                           BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                              SubItemIndex    =   1
                              Text            =   "Date"
                              Object.Width           =   1058
                           EndProperty
                           BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                              SubItemIndex    =   2
                              Text            =   "Size"
                              Object.Width           =   706
                           EndProperty
                        End
                        Begin XpressEditorsLibCtl.dxHyperLinkEdit DxHUpdate 
                           Height          =   420
                           Left            =   1890
                           OleObjectBlob   =   "frmMnuOperations.frx":267A
                           TabIndex        =   53
                           Top             =   0
                           Width           =   1605
                        End
                        Begin VB.Label lblIMSMADataUpdate 
                           Caption         =   "IMSMADataUpdate"
                           Height          =   450
                           Left            =   0
                           TabIndex        =   54
                           Top             =   0
                           Width           =   1890
                           WordWrap        =   -1  'True
                        End
                     End
                  End
               End
               Begin C1SizerLibCtl.C1Elastic elTab1 
                  Height          =   4740
                  Left            =   15
                  TabIndex        =   24
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   3555
                  _cx             =   6271
                  _cy             =   8361
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
                  GridCols        =   1
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"frmMnuOperations.frx":2C9D
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin C1SizerLibCtl.C1Elastic elMAHolder1 
                     Height          =   4740
                     Left            =   0
                     TabIndex        =   25
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   3555
                     _cx             =   6271
                     _cy             =   8361
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
                     _GridInfo       =   $"frmMnuOperations.frx":2CCF
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin TatukGIS_XDK10.XGIS_ControlLegend lgdTest 
                        Height          =   4080
                        Left            =   30
                        TabIndex        =   96
                        Top             =   630
                        Width           =   3495
                        BorderStyle     =   0
                        BeginProperty FontTitle {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        FontColorTitle  =   -16777208
                        BeginProperty FontSubtitle {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   6.75
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        FontColorSubTitle=   -16777208
                        Spacing         =   3
                        ReverseOrder    =   -1  'True
                        Align           =   0
                        Ctl3D           =   -1  'True
                        Color           =   -2147483633
                        Enabled         =   -1  'True
                        Object.Visible         =   -1  'True
                        DoubleBuffered  =   -1  'True
                        AllowMove       =   -1  'True
                        AllowActive     =   -1  'True
                        AllowExpand     =   -1  'True
                        AllowParams     =   -1  'True
                        Mode            =   0
                     End
                     Begin VB.Frame FraSource 
                        Caption         =   "IMSMA Source:"
                        Height          =   600
                        Left            =   30
                        TabIndex        =   26
                        Top             =   30
                        Width           =   3495
                        Begin VB.CheckBox chkSource 
                           Caption         =   "IRAQ Landmine Impact Survey"
                           Enabled         =   0   'False
                           Height          =   285
                           Index           =   4
                           Left            =   135
                           TabIndex        =   27
                           Top             =   225
                           Value           =   1  'Checked
                           Width           =   2940
                        End
                        Begin VB.CheckBox chkSource 
                           Caption         =   "MAG Erbil"
                           Height          =   285
                           Index           =   3
                           Left            =   1665
                           TabIndex        =   28
                           Top             =   450
                           Visible         =   0   'False
                           Width           =   1295
                        End
                        Begin VB.CheckBox chkSource 
                           Caption         =   "RMAC South"
                           Height          =   285
                           Index           =   1
                           Left            =   1665
                           TabIndex        =   29
                           Top             =   225
                           Visible         =   0   'False
                           Width           =   1395
                        End
                        Begin VB.CheckBox chkSource 
                           Caption         =   "IKMAA"
                           Height          =   285
                           Index           =   0
                           Left            =   135
                           TabIndex        =   31
                           Top             =   225
                           Visible         =   0   'False
                           Width           =   1995
                        End
                        Begin VB.CheckBox chkSource 
                           Caption         =   "RMAC Central"
                           Height          =   285
                           Index           =   2
                           Left            =   135
                           TabIndex        =   30
                           Top             =   450
                           Visible         =   0   'False
                           Width           =   1590
                        End
                     End
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic elMineActionNavigator 
               Height          =   480
               Left            =   15
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   4785
               Width           =   3585
               _cx             =   6324
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
               Begin XpressEditorsLibCtl.dxHyperLinkEdit DxHUpdateData 
                  Height          =   540
                  Left            =   1215
                  OleObjectBlob   =   "frmMnuOperations.frx":2D10
                  TabIndex        =   20
                  Top             =   0
                  Width           =   1545
               End
               Begin XpressEditorsLibCtl.dxHyperLinkEdit DxHSettings 
                  Height          =   540
                  Left            =   2745
                  OleObjectBlob   =   "frmMnuOperations.frx":33D7
                  TabIndex        =   21
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1545
               End
               Begin XpressEditorsLibCtl.dxHyperLinkEdit DxHAnalyze 
                  Height          =   540
                  Left            =   45
                  OleObjectBlob   =   "frmMnuOperations.frx":3B15
                  TabIndex        =   22
                  Top             =   0
                  Width           =   1185
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elSecurity 
            Height          =   5280
            Left            =   4350
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   330
            Width           =   3615
            _cx             =   6376
            _cy             =   9313
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
            GridRows        =   3
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmMnuOperations.frx":4373
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin TatukGIS_XDK10.XGIS_ControlLegend LgdSecurity 
               Height          =   3630
               Left            =   0
               TabIndex        =   95
               Top             =   825
               Width           =   3615
               BorderStyle     =   0
               BeginProperty FontTitle {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontColorTitle  =   -16777208
               BeginProperty FontSubtitle {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontColorSubTitle=   -16777208
               Spacing         =   3
               ReverseOrder    =   -1  'True
               Align           =   0
               Ctl3D           =   -1  'True
               Color           =   -2147483633
               Enabled         =   -1  'True
               Object.Visible         =   -1  'True
               DoubleBuffered  =   -1  'True
               AllowMove       =   -1  'True
               AllowActive     =   -1  'True
               AllowExpand     =   -1  'True
               AllowParams     =   -1  'True
               Mode            =   0
            End
            Begin C1SizerLibCtl.C1Elastic elSecurityMeny 
               Height          =   825
               Left            =   0
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   4455
               Width           =   3615
               _cx             =   6376
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
               Begin VB.CommandButton cmdSetScoring 
                  Caption         =   "Set Scoring"
                  Height          =   240
                  Left            =   3480
                  TabIndex        =   71
                  Top             =   540
                  Visible         =   0   'False
                  Width           =   150
               End
               Begin VB.CommandButton cmdSetScope 
                  Height          =   465
                  Left            =   2745
                  MaskColor       =   &H00FFFFFF&
                  Picture         =   "frmMnuOperations.frx":43BC
                  Style           =   1  'Graphical
                  TabIndex        =   55
                  ToolTipText     =   "Set Data Animation"
                  Top             =   90
                  UseMaskColor    =   -1  'True
                  Width           =   480
               End
               Begin VB.Frame FraAnalysisLevel 
                  Caption         =   "Analysis Level:"
                  Height          =   780
                  Left            =   -18000
                  TabIndex        =   40
                  Top             =   780
                  Visible         =   0   'False
                  Width           =   1725
                  Begin VB.CommandButton cmdReset 
                     Caption         =   "Clear"
                     Height          =   255
                     Left            =   1140
                     TabIndex        =   69
                     Top             =   450
                     Width           =   525
                  End
                  Begin VB.OptionButton OptLevel 
                     Caption         =   "cell"
                     Height          =   195
                     Index           =   2
                     Left            =   1080
                     TabIndex        =   43
                     Top             =   225
                     Width           =   555
                  End
                  Begin VB.OptionButton OptLevel 
                     Caption         =   "District"
                     Height          =   195
                     Index           =   1
                     Left            =   45
                     TabIndex        =   42
                     Top             =   495
                     Width           =   855
                  End
                  Begin VB.OptionButton OptLevel 
                     Caption         =   "Province"
                     Height          =   195
                     Index           =   0
                     Left            =   45
                     TabIndex        =   41
                     Top             =   225
                     Value           =   -1  'True
                     Width           =   1095
                  End
               End
               Begin VB.CommandButton cmdSecurityAnalysis 
                  Height          =   465
                  Left            =   -180
                  MaskColor       =   &H00FFFFFF&
                  Picture         =   "frmMnuOperations.frx":489A
                  Style           =   1  'Graphical
                  TabIndex        =   39
                  ToolTipText     =   "Security Analysis"
                  Top             =   1440
                  UseMaskColor    =   -1  'True
                  Width           =   480
               End
               Begin VB.CommandButton cmdInsertIncident 
                  Height          =   465
                  Left            =   3195
                  MaskColor       =   &H00FFFFFF&
                  Picture         =   "frmMnuOperations.frx":4D50
                  Style           =   1  'Graphical
                  TabIndex        =   44
                  ToolTipText     =   "Add New Incident"
                  Top             =   90
                  UseMaskColor    =   -1  'True
                  Width           =   480
               End
            End
            Begin ActiveBar3LibraryCtl.ActiveBar3 abGridTools 
               Height          =   825
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Width           =   3615
               _LayoutVersion  =   2
               _ExtentX        =   6376
               _ExtentY        =   1455
               _DataPath       =   ""
               Bands           =   "frmMnuOperations.frx":5271
            End
         End
         Begin C1SizerLibCtl.C1Elastic elLIS 
            Height          =   5280
            Left            =   4950
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   330
            Width           =   3615
            _cx             =   6376
            _cy             =   9313
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
            BorderWidth     =   1
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
            _GridInfo       =   $"frmMnuOperations.frx":677F
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic elAnalysisTool 
               Height          =   660
               Left            =   15
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   4605
               Width           =   3585
               _cx             =   6324
               _cy             =   1164
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
               Begin VB.CommandButton cmdAnalyse 
                  Caption         =   "Analyze"
                  Height          =   420
                  Left            =   90
                  TabIndex        =   8
                  ToolTipText     =   "Open LIS Analyze tool"
                  Top             =   135
                  Width           =   1365
               End
            End
            Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
               Height          =   4590
               Left            =   15
               OleObjectBlob   =   "frmMnuOperations.frx":67C1
               TabIndex        =   6
               Top             =   15
               Width           =   3585
            End
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   5280
            Left            =   5550
            TabIndex        =   4
            Top             =   330
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   9313
            _Version        =   393217
            Style           =   7
            Appearance      =   1
         End
         Begin C1SizerLibCtl.C1Elastic elLegend 
            Height          =   5280
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   3615
            _cx             =   6376
            _cy             =   9313
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
            GridRows        =   3
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmMnuOperations.frx":7469
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin TatukGIS_XDK10.XGIS_ControlLegend Legend1 
               Height          =   5280
               Left            =   0
               TabIndex        =   97
               Top             =   0
               Width           =   3615
               BorderStyle     =   0
               BeginProperty FontTitle {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontColorTitle  =   -16777208
               BeginProperty FontSubtitle {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontColorSubTitle=   -16777208
               Spacing         =   3
               ReverseOrder    =   0   'False
               Align           =   0
               Ctl3D           =   0   'False
               Color           =   -2147483633
               Enabled         =   -1  'True
               Object.Visible         =   -1  'True
               DoubleBuffered  =   -1  'True
               AllowMove       =   -1  'True
               AllowActive     =   -1  'True
               AllowExpand     =   -1  'True
               AllowParams     =   -1  'True
               Mode            =   0
            End
            Begin C1SizerLibCtl.C1Elastic elThematics 
               Height          =   750
               Left            =   0
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   4530
               Width           =   3615
               _cx             =   6376
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
               Begin VB.PictureBox Picture2 
                  Height          =   615
                  Left            =   1110
                  ScaleHeight     =   555
                  ScaleWidth      =   705
                  TabIndex        =   73
                  Top             =   -660
                  Visible         =   0   'False
                  Width           =   765
               End
               Begin VB.PictureBox Picture1 
                  Height          =   495
                  Left            =   360
                  ScaleHeight     =   435
                  ScaleWidth      =   555
                  TabIndex        =   72
                  Top             =   -840
                  Visible         =   0   'False
                  Width           =   615
               End
               Begin VB.ComboBox ComThematics 
                  Height          =   315
                  Left            =   1260
                  Style           =   2  'Dropdown List
                  TabIndex        =   65
                  Top             =   405
                  Width           =   2355
               End
               Begin VB.ComboBox ComThematicsGroup 
                  Height          =   315
                  Left            =   1260
                  Style           =   2  'Dropdown List
                  TabIndex        =   64
                  Top             =   90
                  Width           =   2355
               End
               Begin VB.Label lblAnalysisType 
                  Caption         =   "Analysis Type:"
                  Height          =   330
                  Left            =   45
                  TabIndex        =   67
                  Top             =   450
                  Width           =   1095
               End
               Begin VB.Label lblAnalysisGroup 
                  Caption         =   "Analysis Group:"
                  Height          =   240
                  Left            =   45
                  TabIndex        =   66
                  Top             =   135
                  Width           =   1140
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elW3 
            Height          =   5280
            Left            =   5250
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   330
            Width           =   3615
            _cx             =   6376
            _cy             =   9313
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
            BorderWidth     =   1
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
            _GridInfo       =   $"frmMnuOperations.frx":74B3
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic elW3Tool 
               Height          =   660
               Left            =   15
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   4605
               Width           =   3585
               _cx             =   6324
               _cy             =   1164
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
               Caption         =   "s"
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
               Begin VB.OptionButton OptW3 
                  Caption         =   "Who"
                  Height          =   195
                  Index           =   0
                  Left            =   2340
                  TabIndex        =   14
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   780
               End
               Begin VB.OptionButton OptW3 
                  Caption         =   "Where"
                  Height          =   240
                  Index           =   2
                  Left            =   2340
                  TabIndex        =   16
                  Top             =   405
                  Visible         =   0   'False
                  Width           =   780
               End
               Begin VB.OptionButton OptW3 
                  Caption         =   "What"
                  Height          =   240
                  Index           =   1
                  Left            =   2340
                  TabIndex        =   15
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   780
               End
               Begin VB.CommandButton cmdCheck 
                  Caption         =   "Check"
                  Height          =   420
                  Left            =   45
                  TabIndex        =   13
                  Top             =   135
                  Width           =   1050
               End
               Begin VB.CommandButton cmdEdit 
                  Caption         =   "Edit"
                  Height          =   420
                  Left            =   1080
                  TabIndex        =   11
                  Top             =   135
                  Width           =   1050
               End
            End
            Begin DXDBGRIDLibCtl.dxDBGrid dxW3DBGrid 
               Height          =   4590
               Left            =   15
               OleObjectBlob   =   "frmMnuOperations.frx":74F5
               TabIndex        =   12
               Top             =   15
               Width           =   3585
            End
         End
         Begin C1SizerLibCtl.C1Tab C1TSearchResults 
            Height          =   5280
            Left            =   6450
            TabIndex        =   74
            Top             =   330
            Width           =   3615
            _cx             =   6376
            _cy             =   9313
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
            Caption         =   "Search|Results"
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
            Begin C1SizerLibCtl.C1Elastic elSearcgResults 
               Height          =   4935
               Left            =   1035
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   330
               Width           =   300
               _cx             =   529
               _cy             =   8705
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
               BorderWidth     =   3
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
               GridRows        =   3
               GridCols        =   5
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmMnuOperations.frx":81B7
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.CommandButton cmdClearSearches 
                  Caption         =   "Clear Searches"
                  Height          =   255
                  Left            =   45
                  TabIndex        =   77
                  Top             =   4425
                  Width           =   195
               End
               Begin VB.CommandButton cmdCancel 
                  Caption         =   "Cancel"
                  Height          =   0
                  Left            =   270
                  TabIndex        =   76
                  Top             =   4425
                  Width           =   0
               End
               Begin vbalTreeViewLib6.vbalTreeView vbalSearch 
                  Height          =   4350
                  Left            =   45
                  TabIndex        =   78
                  Top             =   45
                  Width           =   195
                  _ExtentX        =   344
                  _ExtentY        =   7673
                  Indentation     =   30
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin CONTROLSLibCtl.dxProgressBar dxProgressBar1 
                  Height          =   180
                  Left            =   45
                  TabIndex        =   92
                  Top             =   4710
                  Width           =   195
                  _Version        =   65536
                  _cx             =   344
                  _cy             =   317
                  ForeColor       =   0
                  BackColor       =   14215660
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
                  Pos             =   50
                  Step            =   1
                  ShowText        =   -1  'True
                  Orientation     =   0
                  StartColor      =   16711680
                  EndColor        =   16777215
                  DrawBorderStyle =   1
                  ShowTextStyle   =   1
                  DrawBarStyle    =   2
                  DrawBarBorderStyle=   2
               End
            End
            Begin C1SizerLibCtl.C1Elastic elSrchSettings 
               Height          =   4935
               Left            =   45
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   330
               Width           =   300
               _cx             =   529
               _cy             =   8705
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
               Begin VB.Frame FraSearch 
                  Caption         =   "Settings:"
                  Height          =   3015
                  Left            =   0
                  TabIndex        =   82
                  Top             =   0
                  Width           =   3495
                  Begin VB.Frame FraSearchValue 
                     Caption         =   "Search value:"
                     Height          =   615
                     Left            =   60
                     TabIndex        =   89
                     Top             =   2280
                     Width           =   3375
                     Begin VB.TextBox txtSearchVal 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   90
                        Top             =   240
                        Width           =   3165
                     End
                  End
                  Begin VB.CheckBox chkSearchAll 
                     Caption         =   "Search All Layers (This is slow!)"
                     Height          =   255
                     Left            =   60
                     TabIndex        =   88
                     Top             =   1920
                     Width           =   2595
                  End
                  Begin VB.Frame FraLayerTo 
                     Caption         =   "Layer To Search:"
                     Height          =   1665
                     Left            =   60
                     TabIndex        =   83
                     Top             =   240
                     Width           =   3405
                     Begin VB.CheckBox chkSearchAllFields 
                        Caption         =   "Search All Fields (May Be Slow)"
                        Height          =   285
                        Left            =   60
                        TabIndex        =   87
                        Top             =   1290
                        Width           =   3075
                     End
                     Begin VB.Frame FraFieldTo 
                        Caption         =   "Field To Search:"
                        Height          =   645
                        Left            =   60
                        TabIndex        =   85
                        Top             =   600
                        Width           =   2925
                        Begin VB.ComboBox ComFields 
                           Height          =   315
                           Left            =   60
                           Style           =   2  'Dropdown List
                           TabIndex        =   86
                           Top             =   210
                           Width           =   2775
                        End
                     End
                     Begin VB.ComboBox ComLayer 
                        Height          =   315
                        Left            =   120
                        Style           =   2  'Dropdown List
                        TabIndex        =   84
                        Top             =   240
                        Width           =   2865
                     End
                  End
               End
               Begin VB.Frame Frame2 
                  BorderStyle     =   0  'None
                  Height          =   555
                  Left            =   2220
                  TabIndex        =   80
                  Top             =   3060
                  Width           =   1515
                  Begin VB.CommandButton cmdOK 
                     Caption         =   "Search"
                     Height          =   405
                     Left            =   0
                     TabIndex        =   81
                     Top             =   60
                     Width           =   1275
                  End
               End
            End
            Begin VB.Label lblLblProgress 
               Caption         =   "lblProgress"
               Height          =   4935
               Left            =   1335
               TabIndex        =   91
               Top             =   330
               Width           =   300
            End
         End
      End
   End
   Begin VB.Menu mnuLegendOptions 
      Caption         =   "mnuLegendOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuToggleLegendView 
         Caption         =   "Group / Ungroup"
      End
      Begin VB.Menu mnuSaveMyMap 
         Caption         =   "Save My Map"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuZoomExtent 
         Caption         =   "Zoom to extent"
      End
      Begin VB.Menu mnuLayer 
         Caption         =   "Layer"
         Index           =   0
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Load/Save Settings"
         Begin VB.Menu mnuLoadSettings 
            Caption         =   "Load Settings"
         End
         Begin VB.Menu mnuSaveSettings 
            Caption         =   "Save Settings"
         End
         Begin VB.Menu mnuSaveAllSettings 
            Caption         =   "Save All Settings"
         End
      End
      Begin VB.Menu mnuexportLayer 
         Caption         =   "Export Layer..."
         Begin VB.Menu mnyExportShape 
            Caption         =   "Shape File (ESRI)"
         End
         Begin VB.Menu mnuExportKML 
            Caption         =   "KML File (Google)"
         End
         Begin VB.Menu mnuexportDXF 
            Caption         =   "DXF File (Autocad)"
         End
         Begin VB.Menu mnuExportGPX 
            Caption         =   "GPX File (Open GPS format)"
         End
         Begin VB.Menu mnuExportGML 
            Caption         =   "GML File (Geographic Markup Language)"
         End
         Begin VB.Menu mnuExportTab 
            Caption         =   "TAB File (Mapinfo Native)"
         End
         Begin VB.Menu mnuExportSQLLyr 
            Caption         =   "SQL Layer (Database Layers)"
            Begin VB.Menu mnuExportOgisFormat 
               Caption         =   "Open GIS Format"
            End
            Begin VB.Menu mnuExportNativeFormat 
               Caption         =   "Native Format"
            End
         End
      End
      Begin VB.Menu mnuExportAll 
         Caption         =   "Export All"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExportSelected 
         Caption         =   "Export to new layer from selected"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLegendSettings 
         Caption         =   "Legend/Layer Settings"
         Visible         =   0   'False
         Begin VB.Menu mnuStartGrouped 
            Caption         =   "Start grouped"
         End
         Begin VB.Menu mnuAllowAllowSaveSettings 
            Caption         =   "Allow save settings"
         End
      End
   End
   Begin VB.Menu mnuSearchRight 
      Caption         =   "mnuSearchRight"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveSearch 
         Caption         =   "Remove Search Item"
      End
   End
   Begin VB.Menu mnuGeoItem 
      Caption         =   "mnuGeoItem"
      Visible         =   0   'False
      Begin VB.Menu mnuSearchZoom 
         Caption         =   "Zoom To"
      End
      Begin VB.Menu mnuSearchFlash 
         Caption         =   "Flash"
      End
      Begin VB.Menu mnuSearchRemoveGeo 
         Caption         =   "Remove Search Item"
      End
   End
End
Attribute VB_Name = "frmMnuOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt _
                Lib "gdi32" (ByVal hDestDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal dwRop As Long) As Long
                
Public Event ShowIOM()
Public Event UpdateValues()
Public Event OpenLISAnalyzis()
Public Event ActivateSearchW3Tool()
Public Event LayerActivatedStatus(sLayerName As String, bActivated As Boolean)
Public Event CategorizeBy(scategory As String)
Public Event ActivateSecurityIncidents(bActivated As Boolean)
Public Event DoSecurityAnalysis()
Public Event InsertIncident()
Public Event ZoomToLoc(sName As String, sID As String)
Public Event LoadMap(sMapName As String, oExtent As XGIS_Extent)
Public Event CategorizeIncidents()
Public Event LoadOASISIncidents()
Public Event createChart()
Public Event setScope()
Public Event UpdateNFI(sLocation As String, iNeed As Integer, iDelivery As Integer)
Public Event ActivateTheme(sTheme As String, sThemeGR As String)
Public Event DeActivateTheme(sTheme As String, sThemeGR As String)
Public Event ResetSecurityAnalysis()
Public Event ExportLgd()
Public Event ScoringSettings()
Public Event OpenNewMap(sMAP As String)
Public Event MapPreview(oPic As stdole.StdPicture)
Public Event SaveMyMap()
Public Event LoadLayerToGrid(Index As Integer)
Public Event ZoomLyrExtent(oLyrName As String)
Public Event ExportLayer(oLyrName As String, sFilter As String, sDialogTitle As String, sFileExtention As String, fType As OASIS_GIS_DATA_TYPE)
Public Event SaveAllSettings()
Public Event LoadSetting(sLayer As String)
Public Event SaveSetting(sLayer As String)
Public Event ExportAll()
Public Event ChangeActiveLayer(sName As String)
Public Event ShowMapLibraryDLG(sGUID As String, bInfoOnly As Boolean)
Public Event AddLayerNameToComboDropdown(sLayerName As String)

Private lColor, LColor1

Private mOCN As ADODB.Connection

Private m_oSQLIncLyr As TatukGIS_XDK10.XGIS_LayerSqlAdo
Private m_MASynchURLs(10) As String
Private m_CurrentMap As String
Private m_sTheme As String
Private m_sThemeGR As String
Private bDoNotChangeTab As Boolean
'Private MsXmlHttp As New MSXML2.XMLHTTP

Private m_oShp As TatukGIS_XDK10.XGIS_Shape
Public LyrCol As Collection
Private m_bINIT As Boolean
Private m_OGIS As TatukGIS_XDK10.XGIS_Viewer
Dim m_fntItalic As StdFont
Dim m_bCancel As Boolean
Private m_sSearchCurUID As String
Private m_sSearchCurLyr As String
Private m_sSearchNodeKey As String

Public Sub InitSearch(oGIS As TatukGIS_XDK10.XGIS_Viewer)
        '<EhHeader>
        On Error GoTo InitSearch_Err
        '</EhHeader>

        Dim i As Integer
        Dim sCurrItem As String

100     m_bINIT = True
102     Set m_OGIS = oGIS
104     Set LyrCol = New Collection

106     If ComLayer.ListCount > 0 Then

108         sCurrItem = ComLayer.List(ComLayer.ListIndex)

        End If

110     ComLayer.Clear

112     ComLayer.AddItem "--All--"
114     LyrCol.Add "--All--", "--All--"
        
        On Error Resume Next

116     For i = 0 To oGIS.items.Count - 1

118         If GisUtils.IsInherited(oGIS.items.Item(i), "XGIS_LayerVector") Then
120             LyrCol.Add oGIS.items.Item(i).Name, oGIS.items.Item(i).caption
122             ComLayer.AddItem oGIS.items.Item(i).caption 'Name
            End If

        Next

124     If ComLayer.ListCount > 0 Then

126         If Len(sCurrItem) > 0 Then
128             FindIndexStrEx ComLayer, sCurrItem
            Else
130             FindIndexStrEx ComLayer, "--All--"
            End If

        End If

132     ComFields.Clear
134     ComFields.AddItem "--All--"
136     ComFields.ListIndex = 0

138     chkSearchAll.value = vbUnchecked
140     chkSearchAllFields.value = vbUnchecked
142     txtSearchVal.Text = ""

144     m_bINIT = False

        '<EhFooter>
        Exit Sub

InitSearch_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.InitSearch " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function AddSearchNode(sSearchID As String, _
                               sCaption As String) As cTreeViewNode
    vbalSearch.LabelEdit = True
    Set AddSearchNode = vbalSearch.Nodes.Add(, etvwChild, sSearchID, sCaption, frmSearch.imgl.ItemIndex(28))
    AddSearchNode.Bold = True
    AddSearchNode.ShowPlusMinus = True
End Function

Private Sub cmdCancel_Click()
    m_bCancel = True
End Sub

Private Sub cmdClearSearches_Click()
    PrepareTV
    'lstSearchResult.Clear
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
    Dim sSQL As String
    Dim nodSearch As cTreeViewNode
    Dim nodValue As cTreeViewNode
    Dim nodResult As cTreeViewNode
    Dim k As Integer
    Dim j As Integer
    Dim sDel As String
    Dim oShp9 As TatukGIS_XDK10.XGIS_Shape

    C1TSearchResults.CurrTab = 1

    'lstSearchResult.Clear

    'TODO Strings works only now... Add the Doevents for this to cancel the Search. Do a nice way to present the found features and enable Zoom To!

    '  String = 0
    '  number = 1
    '  Float = 2
    '  Boolean = 3
    '  date = 4

    If chkSearchAllFields.value = vbChecked Then
        dxProgressBar1.Pos = 0
        dxProgressBar1.DoStep
        
        With ComLayer

            If .List(.ListIndex) = "--All--" Then

                For j = 0 To m_OGIS.items.Count - 1

                    If GisUtils.IsInherited(m_OGIS.items.Item(j), "XGIS_LayerVector") Then
                        Set oLyr = m_OGIS.get(m_OGIS.items.Item(j).Name)

                        If oLyr Is Nothing Then Exit Sub

                        DoEvents

                        If m_bCancel Then
                            m_bCancel = False
                            Exit Sub
                        End If

                        Set nodSearch = AddSearchNode(oLyr.caption & ":::" & txtSearchVal.Text & ":::" & "--All--" & ":::" & Now, oLyr.caption & " : " & txtSearchVal.Text)
                        nodSearch.ShowPlusMinus = True
                        nodSearch.Expanded = True
                        
                        For i = 0 To oLyr.Fields.Count - 1

                            If dxProgressBar1.Pos = 100 Then dxProgressBar1.Pos = 1
                            dxProgressBar1.DoStep
        
                            Select Case oLyr.FieldInfo(CLng(i)).FieldType

                                Case TatukGIS_XDK10.XgisFieldTypeString
                                    sDel = "'"

                                Case TatukGIS_XDK10.XgisFieldTypeBoolean
                                    sDel = ""

                                Case TatukGIS_XDK10.XgisFieldTypeDate
                                    sDel = "#"

                                Case TatukGIS_XDK10.XgisFieldTypeFloat, TatukGIS_XDK10.XgisFieldTypeNumber
                                    sDel = ""
                            End Select

                            If oLyr.FieldInfo(CLng(i)).FieldType = TatukGIS_XDK10.XgisFieldTypeString Then

                                For Each oShp9 In oLyr.Loop(oLyr.Extent, oLyr.FieldInfo(CLng(i)).Name & " = '" & txtSearchVal.Text & "'", Nothing, "", True)
                            
                                    'lstSearchResult.AddItem oLyr.Shape.uID
                                    Set nodValue = nodSearch.AddChildNode(oShp9.uID & ":::" & Now() & Rnd(100), "GeoID:" & oShp9.uID, frmSearch.imgl.ItemIndex(97)) ', imgl.ItemIndex(99))
                                    nodValue.ShowPlusMinus = True
                                    nodValue.Tag = oLyr.Name

                                    For k = 0 To oLyr.Fields.Count - 1

                                        Set nodResult = nodValue.AddChildNode(oShp9.uID & ":::" & Now() & Rnd(100), oLyr.Fields.Item(k).Name & " = " & oShp9.GetField(oLyr.Fields.Item(k).Name), frmSearch.imgl.ItemIndex(64))

                                        If i = k Then nodResult.Bold = True
                                        If dxProgressBar1.Pos = 100 Then dxProgressBar1.Pos = 1
                                        dxProgressBar1.DoStep

                                        DoEvents

                                        If m_bCancel Then
                                            m_bCancel = False
                                            Exit Sub
                                        End If

                                    Next
                                Next

                            End If

                        Next

                    End If

                    If Not nodSearch Is Nothing Then
                        If nodSearch.Children.Count = 0 Then
                            vbalSearch.Nodes.Remove nodSearch.Key
                        End If
                    End If

                Next j

            Else
                Set oLyr = m_OGIS.get(LyrCol.Item(ComLayer.List(ComLayer.ListIndex)))
                
                dxProgressBar1.Pos = 0
                
                If oLyr Is Nothing Then Exit Sub
                
                dxProgressBar1.DoStep
                            
                Set nodSearch = AddSearchNode(oLyr.caption & ":::" & txtSearchVal.Text & ":::" & "--All--" & ":::" & Now, oLyr.caption & " : " & txtSearchVal.Text)
                nodSearch.ShowPlusMinus = True
                nodSearch.Expanded = True

                For i = 0 To oLyr.Fields.Count - 1

                    If dxProgressBar1.Pos = 100 Then dxProgressBar1.Pos = 1
                    dxProgressBar1.DoStep

                    If oLyr.FieldInfo(CLng(i)).FieldType = TatukGIS_XDK10.XgisFieldTypeString Then

                        For Each oShp9 In oLyr.Loop(oLyr.Extent, oLyr.FieldInfo(CLng(i)).Name & " = '" & txtSearchVal.Text & "'", Nothing, "", True)
                            
                            Set nodValue = nodSearch.AddChildNode(oShp9.uID & ":::" & Now() & Rnd(100), "GeoID:" & oShp9.uID, frmSearch.imgl.ItemIndex(97)) ', imgl.ItemIndex(99))
                            nodValue.ShowPlusMinus = True
                            nodValue.Tag = oLyr.Name

                            For k = 0 To oLyr.Fields.Count - 1

                                If dxProgressBar1.Pos = 100 Then dxProgressBar1.Pos = 1
                                dxProgressBar1.DoStep
                                Set nodResult = nodValue.AddChildNode(oShp9.uID & ":::" & Now() & Rnd(100), oLyr.Fields.Item(k).Name & " = " & oShp9.GetField(oLyr.Fields.Item(k).Name), frmSearch.imgl.ItemIndex(64))

                                If i = k Then nodResult.Bold = True
                            Next
                        Next

                    End If

                Next

            End If

        End With

    ElseIf chkSearchAll.value = vbChecked Then

    Else

        Set oLyr = m_OGIS.get(LyrCol.Item(ComLayer.List(ComLayer.ListIndex)))
        dxProgressBar1.Pos = 0
                                        
        If oLyr Is Nothing Then Exit Sub

        'Set nodSearch = AddSearchNode(oLyr.caption & ":::" & txtSearchVal.Text & ":::" & Now, oLyr.caption & " : " & txtSearchVal.Text)
        
        If ComFields.List(ComFields.ListIndex) = "--All--" Then
            Set nodSearch = AddSearchNode(oLyr.caption & ":::" & txtSearchVal.Text & ":::" & "--All--" & ":::" & Now, oLyr.caption & " : " & txtSearchVal.Text)
            nodSearch.ShowPlusMinus = True
            nodSearch.Expanded = True

            If oLyr.FieldInfo(oLyr.FindField(ComFields.List(ComFields.ListIndex))).FieldType = TatukGIS_XDK10.XgisFieldTypeString Then

                For k = 0 To oLyr.Field.Count - 1
                    
                    If dxProgressBar1.Pos = 100 Then dxProgressBar1.Pos = 1
                    dxProgressBar1.DoStep
                                        
                    For Each oShp9 In oLyr.Loop(oLyr.Extent, ComFields.List(ComFields.ListIndex) & " = '" & txtSearchVal.Text & "'", Nothing, "", True)
                        'lstSearchResult.AddItem oLyr.Shape.GetField(ComFields.List(ComFields.ListIndex))
                        Set nodValue = nodSearch.AddChildNode(oShp9.uID & ":::" & Now(), "GeoID:" & oShp9.uID, frmSearch.imgl.ItemIndex(97)) ', imgl.ItemIndex(99))
                        nodValue.ShowPlusMinus = True
                        nodValue.Tag = oLyr.Name

                        If dxProgressBar1.Pos = 100 Then dxProgressBar1.Pos = 1
                        dxProgressBar1.DoStep
                                        
                        For i = 0 To oLyr.Field.Count - 1

                            If dxProgressBar1.Pos = 100 Then dxProgressBar1.Pos = 1
                            dxProgressBar1.DoStep
                            nodValue.AddChildNode oShp9.uID & ":::" & Now() & Rnd(100), oLyr.Fields.Item(i).Name & " = " & oShp9.GetField(oLyr.Fields.Item(i).Name), frmSearch.imgl.ItemIndex(64)
                        Next
                    Next

                Next

            End If

        Else
            Set nodSearch = AddSearchNode(oLyr.caption & ":::" & txtSearchVal.Text & ":::" & ComFields.List(ComFields.ListIndex) & ":::" & Now, "Search: " & oLyr.caption & " for " & txtSearchVal.Text & " in " & ComFields.List(ComFields.ListIndex))
            nodSearch.ShowPlusMinus = True
            nodSearch.Expanded = True

            If oLyr.FieldInfo(oLyr.FindField(ComFields.List(ComFields.ListIndex))).FieldType = TatukGIS_XDK10.XgisFieldTypeString Then
                
                For Each oShp9 In oLyr.Loop(oLyr.Extent, ComFields.List(ComFields.ListIndex) & " = '" & txtSearchVal.Text & "'", Nothing, "", True)
                    Set nodValue = nodSearch.AddChildNode(oShp9.uID & ":::" & Now(), "GeoID:" & oShp9.uID, frmSearch.imgl.ItemIndex(97)) ', imgl.ItemIndex(99))
                    nodValue.ShowPlusMinus = True
                    nodValue.Tag = oLyr.Name

                    If dxProgressBar1.Pos = 100 Then dxProgressBar1.Pos = 1
                    dxProgressBar1.DoStep
                    
                    For i = 0 To oLyr.Field.Count - 1

                        If dxProgressBar1.Pos = 100 Then dxProgressBar1.Pos = 1
                        dxProgressBar1.DoStep
                        nodValue.AddChildNode oShp9.uID & ":::" & Now() & Rnd(100), oLyr.Fields.Item(i).Name & " = " & oShp9.GetField(oLyr.Fields.Item(i).Name), frmSearch.imgl.ItemIndex(64)
                    Next

                Next

            End If

        End If

    End If

    dxProgressBar1.Pos = 0
                
End Sub

Private Sub ComLayer_Click()
        '<EhHeader>
        On Error GoTo ComLayer_Click_Err
        '</EhHeader>
        Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
        Dim i As Integer

100     If m_bINIT Then Exit Sub

102    ' DebugPrint

104     Set oLyr = m_OGIS.get(LyrCol.Item(ComLayer.List(ComLayer.ListIndex)))

106     If oLyr Is Nothing Then Exit Sub

108     With oLyr.Fields

110         ComFields.Clear

112         For i = 0 To .Count - 1
114             ComFields.AddItem .Item(i).Name
                DebugPrint .Item(i).Name
                DebugPrint oLyr.FieldInfo(i).Binary
                DebugPrint oLyr.FieldInfo(i).Decimal
                DebugPrint oLyr.FieldInfo(i).Deleted
                DebugPrint oLyr.FieldInfo(i).ExportName
                DebugPrint oLyr.FieldInfo(i).FieldType
                DebugPrint oLyr.FieldInfo(i).FileFormat
                DebugPrint oLyr.FieldInfo(i).Hidden
                DebugPrint oLyr.FieldInfo(i).Predefied
                DebugPrint oLyr.FieldInfo(i).Width
            Next

        End With

        '<EhFooter>
        Exit Sub

ComLayer_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmSearch.ComLayer_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub PrepareTV()

    With vbalSearch
        .ImageList = frmSearch.imgl.hIml
        .Nodes.Clear
    End With

    Set m_fntItalic = New StdFont
    m_fntItalic.Name = "Tahoma"
    m_fntItalic.Size = 8.25
    'm_fntItalic.Italic = True
    m_fntItalic.Underline = True
    m_fntItalic.Bold = True

End Sub

Private Sub ctlMapFileContainer1_MapDelete(sGUID As String)

    If MsgBox("Are you sure you want to delete to active map?", vbYesNo) = vbYes Then
        m_Cnn.Execute "delete FROM [ttkGISProjectDef] where sGUID = '" & sGUID & "'"
        ctlMapFileContainer1.Init
        
        m_OGIS.UpDate
    End If
        
End Sub

Private Sub ctlMapFileContainer1_MapInfo(sInfo As String, _
                                         sGUID As String)
    RaiseEvent ShowMapLibraryDLG(sGUID, True) ' MsgBox "show map info here"
End Sub

Private Sub ctlMapFileContainer1_MapLoad(sPath As String, _
                                         oExtent As TatukGIS_XDK10.XGIS_Extent)
    Me.Enabled = False
    bDoNotChangeTab = True
    RaiseEvent LoadMap(sPath, oExtent)
    bDoNotChangeTab = False
    Me.Enabled = True
    
    'C1TabOpsView.CurrTab = 6
End Sub

Public Sub NewMapLibraryItem(sTitle As String, _
                             sInfo As String, _
                             sCreatedBy As String, _
                             sCopyright As String, _
                             sURL As String, _
                             sSource As String, _
                             dXMin As String, _
                             dXMax As String, _
                             dYMin As String, _
                             dYMax As String, _
                             dcenterX As String, _
                             dcenterY As String, _
                             sscale As String, _
                             sEPSG As String, _
                             sCreatedDate As String)
        '<EhHeader>
        On Error GoTo NewMapLibraryItem_Err
        '</EhHeader>
                                                                                                
        Dim oStream As ADODB.Stream
        Dim oRS As ADODB.Recordset
        Dim oPic As Picture
        Dim sPath As String
        Dim sGUID As String
        Dim sMapData As String
        Dim oPB As PictureBox
        Dim oRect As New XRect
    
100     Set oRS = New ADODB.Recordset
        
102     With oRS
104         ctlMapFileContainer1.SetCurrentActiveMapAsInactive
106         .Open "SELECT * FROM [ttkGISProjectDef] where sGUID = 'Bad DoyleSolutions'", m_Cnn, adOpenDynamic, adLockBatchOptimistic
108         .AddNew
            
110         sGUID = GUIDGen
112         .Fields("sGUID").value = sGUID
114         .Fields("InUse").value = False
116         .Fields("XMIN").value = dXMin
118         .Fields("XMAX").value = dXMax
120         .Fields("YMIN").value = dYMin
122         .Fields("YMAX").value = dYMax
124         .Fields("sName").value = sTitle
126         .Fields("sInfo").value = sInfo
128         .Fields("bSavedToDB").value = True
130         .Fields("sFilePath").value = ""
132         .Fields.Item("centerX").value = dcenterX
134         .Fields.Item("centerY").value = dcenterY
136         .Fields.Item("scale").value = sscale
138         .Fields.Item("EPSG").value = sEPSG
140         .Fields.Item("CreatedBy").value = sCreatedBy
142         .Fields.Item("Copyright").value = sCopyright
144         .Fields.Item("url").value = sURL
146         .Fields.Item("Source").value = sSource
148         .Fields.Item("CreatedDate").value = sCreatedDate
150         sPath = g_sAppPath & "\data\user\maps\" & sGUID & ".ttkgp"
        
            Dim oExtent As XGIS_Extent
152         Set oExtent = New XGIS_Extent
154         Set oExtent = m_OGIS.VisibleExtent
            On Error Resume Next
156         m_OGIS.Delete m_oSQLIncLyr.Name
158         m_OGIS.Delete "Draw_Layer"
            On Error GoTo NewMapLibraryItem_Err
162         m_OGIS.SaveProjectAs sPath, False
160         m_OGIS.SaveAll
            
164         Set oStream = New ADODB.Stream
166         oStream.Open
168         oStream.Type = 2
170         oStream.Charset = "ascii"
172         oStream.LoadFromFile sPath

174         sMapData = oStream.ReadText
176         sMapData = Replace$(sMapData, "\\", "\")
178         sMapData = Replace$(sMapData, "\\", "\")
179         sMapData = Replace$(sMapData, "\\", "\")
180         sMapData = Replace$(sMapData, g_sAppPath, "CLIENTDBPATH", , , vbTextCompare)

182         .Fields("MapData").value = sMapData
184         oStream.Close
186         Set oStream = Nothing

188         ctlMapFileContainer1.picPicture.AutoRedraw = True
190         Set ctlMapFileContainer1.picPicture.Picture = Nothing
192         ctlMapFileContainer1.picPicture.Cls
194         Set oPB = ctlMapFileContainer1.picPicture

196         With oPB
198             .Move .left, .top, ScaleX(m_OGIS.ClientWidth, vbPixels, vbTwips) / 4, ScaleY(m_OGIS.ClientHeight, vbPixels, vbTwips) / 4
200             ctlMapFileContainer1.SetExportImageSize oPB
202             oRect.Prepare 0, 0, ScaleX(.Width, vbTwips, vbPixels), ScaleY(.Height, vbTwips, vbPixels)
            End With
       
204         m_OGIS.PrintDC oPB.hdc, 96, oRect, m_OGIS.VisibleExtent, 0
206         'm_OGIS.draw 'this command screws up rendering
208         oPB.Picture = oPB.Image
            
210         ctlMapFileContainer1.AddControl sGUID, -1, -1, sTitle, sPath, sInfo, Me.BackColor, vbBlack, UserCustom, oExtent, oPB     'oPic
212         ctlMapFileContainer1.SaveImageToDB oPB, oRS, "oImagePreview"
214         ctlMapFileContainer1.HighlightMapAsActive -1, sGUID, sPath
216         ctlMapFileContainer1.SetActiveMapAsUserMap
218         .UpdateBatch
220         .Close
222         m_OGIS.UpDate
        End With
    
224     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

NewMapLibraryItem_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.NewMapLibraryItem " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ctlMapFileContainer1_MapNew()
    RaiseEvent ShowMapLibraryDLG("", False)
End Sub

Private Sub ctlMapFileContainer1_MapPreview(oPicture As stdole.StdPicture)
    RaiseEvent MapPreview(oPicture)
    'I need to talk a bit more on this... Shall we trickle down or do things here?
    MsgBox "show map preview here"

End Sub

Public Sub MapLibSaveMap(sPath As String, _
                         sGUID As String, _
                         sTitle As String, _
                         sInfo As String, _
                         sCreatedBy As String, _
                         sCopyright As String, _
                         sURL As String, _
                         sSource As String, _
                         sCreatedDate As String)
        '<EhHeader>
        On Error GoTo MapLibSaveMap_Err
        '</EhHeader>
    
        Dim oRS As ADODB.Recordset
        Dim oPic As Picture
        Dim oRect As New XRect
        Dim ostdPct As StdPicture
        Dim oPB As PictureBox
        Dim oStream As ADODB.Stream
        Dim sMapData As String
        
100     Set oRS = New ADODB.Recordset
102     oRS.Open "SELECT * FROM [ttkGISProjectDef] where sGUID  = '" & sGUID & "'", m_Cnn, adOpenDynamic, adLockBatchOptimistic

104     If Not oRS.EOF Then
       
106         ctlMapFileContainer1.picPicture.AutoRedraw = True
108         Set ctlMapFileContainer1.picPicture.Picture = Nothing
110         ctlMapFileContainer1.picPicture.Cls
112         Set oPB = ctlMapFileContainer1.picPicture
        
114         With oPB
116             .Move .left, .top, ScaleX(m_OGIS.ClientWidth, vbPixels, vbTwips) / 4, ScaleY(m_OGIS.ClientHeight, vbPixels, vbTwips) / 4
118             ctlMapFileContainer1.SetExportImageSize oPB
120             oRect.Prepare 0, 0, ScaleX(.Width, vbTwips, vbPixels), ScaleY(.Height, vbTwips, vbPixels)
            End With
    
122         m_OGIS.PrintDC oPB.hdc, 96, oRect, m_OGIS.VisibleExtent, 0
124         'm_OGIS.draw 'this command screws up rendering
126         oPB.Picture = oPB.Image

128         ctlMapFileContainer1.SaveActiveMap oPB, sInfo, sTitle, sGUID
            On Error Resume Next
            m_OGIS.Delete m_oSQLIncLyr.Name
            m_OGIS.Delete "Draw_Layer"
            
130         '
132         m_OGIS.SaveProjectAs sPath, False
            m_OGIS.SaveAll
            On Error GoTo MapLibSaveMap_Err
            
134         Set oStream = New ADODB.Stream
136         oStream.Open
138         oStream.Type = 2    ' type = text
140         oStream.Charset = "ascii"
142         oStream.LoadFromFile sPath

            sMapData = oStream.ReadText
            sMapData = Replace$(sMapData, "\\", "\")
            sMapData = Replace$(sMapData, "\\", "\")
            sMapData = Replace$(sMapData, "\\", "\")
            sMapData = Replace$(sMapData, g_sAppPath, "CLIENTDBPATH", , , vbTextCompare)
 
144         oRS.Fields("MapData").value = sMapData
146         oStream.Close
148         Set oStream = Nothing
        
150         With oRS
152             .Fields("XMax").value = m_OGIS.VisibleExtent.xmax
154             .Fields("XMin").value = m_OGIS.VisibleExtent.xmin
156             .Fields("YMax").value = m_OGIS.VisibleExtent.ymax
158             .Fields("YMin").value = m_OGIS.VisibleExtent.ymin
160             .Fields("sName").value = sTitle
162             .Fields("sInfo").value = sInfo
164             .Fields("bSavedToDB").value = True
166             .Fields.Item("centerX").value = m_OGIS.CenterPtg.X
168             .Fields.Item("centerY").value = m_OGIS.CenterPtg.Y
170             .Fields.Item("scale").value = m_OGIS.ScaleAsText
172             .Fields.Item("EPSG").value = m_OGIS.CS.EPSG
174             .Fields.Item("CreatedBy").value = sCreatedBy
176             .Fields.Item("Copyright").value = sCopyright
178             .Fields.Item("url").value = sURL
180             .Fields.Item("Source").value = sSource
182             .Fields.Item("CreatedDate").value = sCreatedDate
            End With

184         oRS.UpdateBatch adAffectCurrent
        
186         ctlMapFileContainer1.SetCurrentExtent m_OGIS.VisibleExtent
188         oRS.Close
190         Set oRS = Nothing
        m_OGIS.UpDate
192         MsgBox "Map saved", vbInformation, "OASIS Map Library"
        
        Else
194         MsgBox "Map save failed", vbCritical, "OASIS Map Library"
        End If

        '<EhFooter>
        Exit Sub

MapLibSaveMap_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.MapLibSaveMap " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ctlMapFileContainer1_MapSave(sPath As String, _
                                         sGUID As String)

100     'If MsgBox("Are you sure you want to save the active map [" & ctlMapFileContainer1.GetActiveMapName & "]?" & vbLf & vbLf & "This operation will override any settings for this map existing.", vbYesNo) = vbYes Then
        RaiseEvent ShowMapLibraryDLG(sGUID, False)
        ' End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_OGIS = Nothing
End Sub

Private Sub Legend1_OnLayerActiveChange(translated As Boolean, _
                                        ByVal layer As Object)
    'RaiseEvent ChangeActiveLayer(Layer.Name)
End Sub

Private Sub Legend1_OnLayerSelect(translated As Boolean, _
                                  ByVal layer As Object)

    'If Not layer Is Nothing Then RaiseEvent ChangeActiveLayer(layer.Name)
End Sub

Private Sub Legend1_OnMouseDown(translated As Boolean, _
                                ByVal Button As TatukGIS_XDK10.XMouseButton, _
                                ByVal Shift As TatukGIS_XDK10.XShiftState, _
                                ByVal X As Long, _
                                ByVal Y As Long)
        '<EhHeader>
        On Error GoTo Legend1_OnMouseDown_Err
        '</EhHeader>
        Dim i As Integer

100     If Button = XmbRight Then
            
            mnuZoomExtent.Visible = False
            mnuSettings.Visible = False
            
102         If Not Legend1.GIS_Layer Is Nothing Then
                
                ' Legend1.GIS_Layer.r
                                
                mnuSettings.Visible = True
                mnuZoomExtent.Visible = True
                
                If GisUtils.IsInherited(Legend1.GIS_Layer, "XGIS_LayerVector") Or GisUtils.IsInherited(Legend1.GIS_Layer, "XGIS_LayerSqlAbstract") Then
                    Debug.Print Legend1.GIS_Layer.caption & " Load to data view"

104                 For i = 1 To mnuLayer.UBound
                        Debug.Print mnuLayer(i).caption
                        
106                     If Legend1.GIS_Layer.caption & " Load to data view" = mnuLayer(i).caption Then
108                         mnuLayer(i).Visible = True
                        Else
110                         mnuLayer(i).Visible = False
                        End If

                    Next

                    mnuexportLayer.caption = "Export Layer:" & Legend1.GIS_Layer.caption
                    mnuexportLayer.Visible = True
                    
                    If Legend1.GIS_Layer.GetSelectedCount > 0 Then
                        mnuExportSelected.Visible = True
                    Else
                        mnuExportSelected.Visible = False
                    End If
                
                End If
                
            Else
                
                mnuexportLayer.Visible = False

112             For i = 1 To mnuLayer.UBound
114                 mnuLayer(i).Visible = False
                Next

            End If
        
116         PopupMenu mnuLegendOptions
    
        End If

        '<EhFooter>
        Exit Sub

Legend1_OnMouseDown_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.Legend1_OnMouseDown " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuexportDXF_Click()
    Dim ef As ExportFile

    With ef
        .sFilter = "Autocad DXF files (*.dxf)|*.dxf"
        .sDialogTitle = "Export to Autocad DXF file"
        .sFileExtention = ".dxf"
        .fType = vDXF
        RaiseEvent ExportLayer(Legend1.GIS_Layer.Name, .sFileExtention, .sDialogTitle, .sFileExtention, .fType)
    End With

End Sub

Private Sub mnuExportGML_Click()
    Dim ef As ExportFile

    With ef
        .sFilter = "OGIS GML files (*.GML)|*.GML"
        .sDialogTitle = "Export to OGIS GML  file"
        .sFileExtention = ".gml"
        .fType = vGML
        RaiseEvent ExportLayer(Legend1.GIS_Layer.Name, .sFileExtention, .sDialogTitle, .sFileExtention, .fType)
    End With

End Sub

Private Sub mnuExportGPX_Click()
    Dim ef As ExportFile

    With ef
        .sFilter = "GPS Export files (*.gpx)|*.gpx"
        .sDialogTitle = .sFilter '.Filename & ".gpx"
        .sFileExtention = ".gpx"
        .fType = vGPX
        RaiseEvent ExportLayer(Legend1.GIS_Layer.Name, .sFileExtention, .sDialogTitle, .sFileExtention, .fType)
    End With

End Sub

Private Sub mnuExportKML_Click()
    Dim ef As ExportFile

    With ef
        .sFilter = "GOOGLE KML files (*.kml)|*.kml"
        .sDialogTitle = "Export to Google KML file"
        .sFileExtention = ".kml"
        .fType = vKML
        RaiseEvent ExportLayer(Legend1.GIS_Layer.Name, .sFileExtention, .sDialogTitle, .sFileExtention, .fType)
    End With
    
End Sub

Private Sub mnuExportNativeFormat_Click()
    ExportVectorToDB False
End Sub

Private Sub mnuExportOgisFormat_Click()
    ExportVectorToDB True
End Sub

Private Sub ExportVectorToDB(bOGISFormat As Boolean)
        '<EhHeader>
        On Error GoTo ExportVectorToDB_Err
        '</EhHeader>
        Dim oLyr As New TatukGIS_XDK10.XGIS_LayerVector
        Dim xportLyr As Object
        Dim c As New cCommonDialog
        Dim sAdoString As String
        Dim sLayerName As String
        Dim oSQLLyr As New TatukGIS_XDK10.XGIS_LayerSqlAdo
        Dim oRS As ADODB.Recordset
        Dim oExtent As TatukGIS_XDK10.XGIS_Extent
        Dim oAbsLayer As XGIS_LayerAbstract
         
100     'frmLayerSelection.Init m_OGIS
102     'frmLayerSelection.Show vbModal, Me

104     If Legend1.GIS_Layer.Name <> "" Then
                    
106         Set oLyr = m_OGIS.get((Legend1.GIS_Layer.Name))

108         With c
110             .Filter = "Microsoft Access (*.mdb)|*.mdb"
112             .DialogTitle = "Create OASIS SQL Layers"
114             .InitDir = g_sAppPath & "\data\db\"
116             .ShowOpen
            End With
                    
118         If Len(c.Filename) > 0 Then
                If bOGISFormat Then
                   Set xportLyr = New TatukGIS_XDK10.XGIS_LayerSqlOgisAdo
                Else
120                 Set xportLyr = New TatukGIS_XDK10.XGIS_LayerSqlAdo
                End If
                
122             sLayerName = InputBox("Specify the new name of the layer (with no spaces please)", "OASIS GIS SQL Layers", "MyLayer")
                sLayerName = Replace$(sLayerName, " ", "")
                
124             If sLayerName = "" Then
126                 MsgBox "It seems like you have not entered a proper name of the table to be created. please try again!", vbInformation, "OASIS Data Creator"
128                 frmLayerSelection.ClearItem
                    Exit Sub
                End If
                        
130             'If MsgBox("Is this the OASIS database?", vbYesNo, "Confirm if this is the OASIS database") = vbYes Then
132                 'sAdoString = GetConnectionString(c.Filename)
                'Else
134                 sAdoString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & c.Filename & ";" 'Mode=ReadWrite|Share Deny None;Persist Security Info=False"
                'End If
                
                If bOGISFormat Then
                    frmOGISFormats.Show vbModal, Me
                    sAdoString = "[TatukGIS Layer]" & vbCrLf & "Storage = " & g_sOGISFormat & vbCrLf & "layer=" & sLayerName & vbCrLf & "Dialect=MSJET" & vbCrLf & "ADO=" & sAdoString
                Else
136                 sAdoString = "[TatukGIS Layer]" & vbCrLf & "Storage = Native" & vbCrLf & "layer=" & sLayerName & vbCrLf & "Dialect=MSJET" & vbCrLf & "ADO=" & sAdoString
                End If
                
138             xportLyr.Path = sAdoString
140             oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
142             xportLyr.SaveAll

                
'                Dim mFileSysObj As New FileSystemObject
'                Dim sPath As String
'
'144             sPath = Mid$(c.Filename, 1, InStrRev(c.Filename, "\")) & sLayerName & ".ttkls"
'
'146             mFileSysObj.CreateTextFile sPath, True
'
'148             Open sPath For Output As #1
'150             Print #1, sAdoString
'152             Close #1
                    
            End If
        End If
                
154     frmLayerSelection.ClearItem

        '<EhFooter>
        Exit Sub

ExportVectorToDB_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.ExportVectorToDB " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuExportSelected_Click()
    Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
    Dim oTargetLyr As TatukGIS_XDK10.XGIS_LayerVector

    Set oLyr = Legend1.GIS_Layer
    Set oTargetLyr = New TatukGIS_XDK10.XGIS_LayerVector
    
    oLyr.ExportLayer oTargetLyr, oLyr.Extent, XgisShapeTypeUnknown, "GIS_SELECTED=True", False
    oTargetLyr.caption = oLyr.caption & " export " & RFC3339DateTime
    oTargetLyr.Name = oTargetLyr.caption
    oTargetLyr.DeselectAll
    oTargetLyr.Tag = 666
    m_OGIS.Add oTargetLyr
    RaiseEvent AddLayerNameToComboDropdown(oTargetLyr.caption)
End Sub

Private Sub mnuExportTab_Click()
    Dim ef As ExportFile

    With ef
        .sFilter = "ESRI Shape files (*.tab)|*.tab"
        .sDialogTitle = "Export to MapInfo TAB file"
        .sFileExtention = ".tab"
        .fType = vTAB
        RaiseEvent ExportLayer(Legend1.GIS_Layer.Name, .sFileExtention, .sDialogTitle, .sFileExtention, .fType)
    End With

End Sub

Private Sub mnuLayer_Click(Index As Integer)
    RaiseEvent LoadLayerToGrid(Index)
    mnuLayer(Index).Visible = False
End Sub

Private Sub mnuLoadSettings_Click()
    
    RaiseEvent LoadSetting(Legend1.GIS_Layer.Name)

End Sub

Private Sub mnuRemoveSearch_Click()

    If m_sSearchNodeKey = "" Then Exit Sub
    vbalSearch.Nodes.Remove m_sSearchNodeKey
    m_sSearchNodeKey = ""
End Sub

Private Sub mnuSaveAllSettings_Click()

    RaiseEvent SaveAllSettings

End Sub

Private Sub mnuSaveMyMap_Click()
    MsgBox "dude"
    
End Sub

Private Sub mnuSaveSettings_Click()
    RaiseEvent SaveSetting(Legend1.GIS_Layer.Name)
End Sub

Private Sub mnuSearchFlash_Click()

    If m_sSearchCurUID = "" Then Exit Sub

    Dim lShape As TatukGIS_XDK10.XGIS_Shape

    Set lShape = m_OGIS.get(m_sSearchCurLyr).GetShape(CLng(m_sSearchCurUID))
    lShape.Flash

    m_sSearchCurLyr = ""
    m_sSearchCurUID = ""

    Set lShape = Nothing
End Sub

Private Sub mnuSearchRemoveGeo_Click()

    If m_sSearchNodeKey = "" Then Exit Sub
    vbalSearch.Nodes.Remove m_sSearchNodeKey
    m_sSearchNodeKey = ""
End Sub

Private Sub mnuSearchZoom_Click()

    If m_sSearchCurUID = "" Then Exit Sub

    Dim i As Integer
    Dim lShape As TatukGIS_XDK10.XGIS_Shape

    Set lShape = m_OGIS.get(m_sSearchCurLyr).GetShape(CLng(m_sSearchCurUID))

    m_OGIS.Lock
    m_OGIS.VisibleExtent = lShape.Extent
    m_OGIS.Unlock

    m_sSearchCurLyr = ""
    m_sSearchCurUID = ""

    Set lShape = Nothing
End Sub

Private Sub mnuToggleLegendView_Click()

    If Legend1.Mode = XgisControlLegendModeGroups Then
        Legend1.Mode = XgisControlLegendModeLayers
    Else
        Legend1.Mode = XgisControlLegendModeGroups
    End If

End Sub

Private Sub mnuZoomExtent_Click()
        '<EhHeader>
        On Error GoTo mnuZoomExtent_Click_Err
        '</EhHeader>

100     RaiseEvent ZoomLyrExtent(Legend1.GIS_Layer.Name)
    
        '<EhFooter>
        Exit Sub

mnuZoomExtent_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.mnuZoomExtent_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnyExportShape_Click()
    Dim ef As ExportFile

    With ef
        .sFilter = "ESRI Shape files (*.shp)|*.shp"
        .sDialogTitle = "Export to ESRI shape file"
        .sFileExtention = ".shp"
        .fType = vSHP
        RaiseEvent ExportLayer(Legend1.GIS_Layer.Name, .sFileExtention, .sDialogTitle, .sFileExtention, .fType)
    End With
    
End Sub

Private Sub vbalSearch_NodeRightClick(Node As vbalTreeViewLib6.cTreeViewNode)
    Dim tP As POINTAPI
    Dim sKey() As String

    If Node.Children.Count = 0 Then Exit Sub 'Attribute field

    m_sSearchNodeKey = Node.Key

    GetCursorPos tP
    ScreenToClient vbalSearch.hwnd, tP

    If Node.Parent Is Nothing Then
        DebugPrint "Root"
        Me.PopupMenu mnuSearchRight, , vbalSearch.left + tP.X * Screen.TwipsPerPixelX, vbalSearch.top + tP.Y * Screen.TwipsPerPixelY

    ElseIf Node.Parent.Parent Is Nothing Then
        DebugPrint "Geoitem"
        m_sSearchCurLyr = Node.Tag
        sKey = Split(Node.Key, ":::")
        m_sSearchCurUID = sKey(0)
        Me.PopupMenu mnuGeoItem, , vbalSearch.left + tP.X * Screen.TwipsPerPixelX, vbalSearch.top + tP.Y * Screen.TwipsPerPixelY

    End If

End Sub

Public Property Get CurrentMap() As String
    CurrentMap = m_CurrentMap
End Property

Public Property Get SecAnalysisevel() As Integer

    If OptLevel(0).value Then
        SecAnalysisevel = 0
    ElseIf OptLevel(1).value Then
        SecAnalysisevel = 1
    Else
        SecAnalysisevel = 2
    End If

End Property

Public Sub refreshGrid()
    dxDBGrid1.Dataset.ADODataset.Requery
End Sub

Private Sub ZoomTofeature(sVals As String)
    '    MsgBox sVals
End Sub

Private Sub abGridTools_ComboSelChange(ByVal Tool As ActiveBar3LibraryCtl.Tool)
        '<EhHeader>
        On Error GoTo abGridTools_ComboSelChange_Err

        '</EhHeader>
100     Select Case Tool.Name

            Case Is = "comIncCategory"
                RaiseEvent CategorizeIncidents
                CategorizeIncidentsByType
                
                LgdSecurity.UpDate

            Case Is = "comDateFrom"
                RaiseEvent LoadOASISIncidents

            Case Else
            
        End Select
    
        '<EhFooter>
        Exit Sub

abGridTools_ComboSelChange_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.abGridTools_ComboSelChange " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub C1TabOpsView_Switch(OldTab As Integer, _
                                NewTab As Integer, _
                                Cancel As Integer)
    
    cmdExportLegend.Visible = IIf(NewTab < 2, True, False)
    
    If NewTab < 2 Then
        cmdExportLegend.Move Me.Width - (cmdExportLegend.Width + CInt(cmdExportLegend.Width / 2) + 70), IIf(NewTab = 0, 368, 1195)
    End If

End Sub

Private Sub chkSource_Click(Index As Integer)
    Dim oLyr As New TatukGIS_XDK10.XGIS_LayerVector
    Dim iIniFile As New TatukGIS_XDK10.XGIS_Ini
    
    Select Case Index
    
        Case 0
            
        Case 1
        
        Case 2
        
        Case 3
        
        Case 4
    
    End Select

End Sub

Private Sub cmdAnalyse_Click()
    RaiseEvent OpenLISAnalyzis
End Sub

Private Sub cmdCommand1_Click()
    lgdTest.UpDate
End Sub

Private Sub cmdExportLegend_Click()
        '<EhHeader>
        On Error GoTo cmdExportLegend_Click_Err
        '</EhHeader>
    
        Dim picStdPicture As New StdPicture
        Dim picOutput As New StdPicture
    
        Dim lTotalHeight As Long
        Dim lControlHeight As Long
        Dim dNumberOfIterations As Double
        Dim i As Integer
        Dim T As Long
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' HIDE THE UNSELECTED LAYERS
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If 1 = 2 Then
            RaiseEvent InsertIncident
        End If
   
100     If C1TabOpsView.CurrTab = 0 Then
        
            Dim l As Long
            Dim b2 As Boolean
            Dim b3 As Boolean
            Dim oLyrAB As TatukGIS_XDK10.XGIS_LayerAbstract
            Dim bLayerExists As Boolean
            Dim sLayerNames() As String
            Dim bLayerSeleted() As Boolean
        
102         l = 0
104         ReDim sLayerNames(l + 1)
106         ReDim bLayerSeleted(l + 1)
        
108         bLayerExists = False
            On Error Resume Next
110         bLayerExists = Legend1.GetLayerInfo(l, sLayerNames(l), bLayerSeleted(l), b2, b3)
            On Error GoTo cmdExportLegend_Click_Err
        
112         Do Until Not bLayerExists
                On Error Resume Next ' new from 03 sept 2013
114             If Not bLayerSeleted(l) Then
            
116                 Set oLyrAB = Legend1.GIS_Viewer.get(sLayerNames(l))
118                 oLyrAB.HideFromLegend = True

                End If
        
120             l = l + 1
122             ReDim Preserve sLayerNames(l + 1)
124             ReDim Preserve bLayerSeleted(l + 1)
            
126             bLayerExists = False
                On Error Resume Next
128             bLayerExists = Legend1.GetLayerInfo(l, sLayerNames(l), bLayerSeleted(l), b2, b3)
                On Error GoTo cmdExportLegend_Click_Err
            Loop
        
130         Legend1.UpDate
        
        End If
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
132     Picture1.AutoSize = True
134     Picture2.AutoSize = True
136     Picture1.AutoRedraw = True
138     Picture2.AutoRedraw = True
140     Clipboard.Clear

142     If C1TabOpsView.CurrTab = 0 Then
144         Legend1.Width = Legend1.Width * 2
146         Legend1.PrintClipboard
        Else
148         LgdSecurity.Width = LgdSecurity.Width * 2
150         LgdSecurity.PrintClipboard
        End If
    
152     Set picOutput = Clipboard.GetData(vbCFMetafile)
154     lTotalHeight = picOutput.Height / Screen.TwipsPerPixelY
156     lTotalHeight = lTotalHeight / 2.35
    
158     If C1TabOpsView.CurrTab = 0 Then
160         lControlHeight = Legend1.Height / Screen.TwipsPerPixelY
        Else
162         lControlHeight = LgdSecurity.Height / Screen.TwipsPerPixelY
        End If
    
164     dNumberOfIterations = lTotalHeight / lControlHeight
166     DebugPrint "Total (" & lTotalHeight & ") / Small (" & lControlHeight & ") = " & dNumberOfIterations
168     Picture1.Height = picOutput.Height
170     Picture1.Width = picOutput.Width
172     Picture1.PaintPicture picOutput, 0, 0

174     T = 0

176     Do Until i > dNumberOfIterations

178         If C1TabOpsView.CurrTab = 0 Then
180             Legend1.PrintBmp picStdPicture, 0, lControlHeight, T, 1
            Else
182             LgdSecurity.PrintBmp picStdPicture, 0, lControlHeight, T, 1
            End If
        
184         Picture2.Height = picStdPicture.Height
186         Picture2.Width = picStdPicture.Width
188         Picture2.PaintPicture picStdPicture, 0, 0
190         BitBlt Picture1.hdc, 0, -T, picStdPicture.Width, picStdPicture.Height, Picture2.hdc, 0, 0, vbSrcCopy
192         T = T - lControlHeight
194         i = i + 1
        Loop
   
196     Picture2.Cls
198     Picture2.Height = lTotalHeight * Screen.TwipsPerPixelY
200     Picture2.Width = Picture2.Width * 0.5
202     BitBlt Picture2.hdc, 0, 0, Picture2.Width, Picture2.Height, Picture1.hdc, 0, 0, vbSrcCopy
     
204     If C1TabOpsView.CurrTab = 0 Then
206         Legend1.Width = Legend1.Width / 2
208         Legend1.PrintBmp picStdPicture, 0, 0, 0, 1
        Else
210         LgdSecurity.Width = LgdSecurity.Width / 2
212         LgdSecurity.PrintBmp picStdPicture, 0, 0, 0, 1
        End If
    
214     Clipboard.SetData Picture2.Image
216     MsgBox "The Legend has been added to your clipboard.", vbInformation, "OASIS Client Export..."
    
218     Picture2.Cls
220     Picture1.Cls
222     Set picStdPicture = Nothing
224     Set picOutput = Nothing

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' unHIDE THE UNSELECTED LAYERS
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
226     If C1TabOpsView.CurrTab = 0 Then
        
228         l = 0

230         Do Until l = UBound(sLayerNames) - 1
        
232             If Not bLayerSeleted(l) Then
            
234                 Set oLyrAB = Legend1.GIS_Viewer.get(sLayerNames(l))
236                 oLyrAB.HideFromLegend = False

                End If
        
238             l = l + 1
            Loop
        
240         Legend1.UpDate
242         Set oLyrAB = Nothing
        End If
    
        '<EhFooter>
        Exit Sub

cmdExportLegend_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.cmdExportLegend_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdInfo_Click(Index As Integer)

    Select Case Index
    
        Case 0

            ' frmMapProductInfo.txtInfo.Text = txtMapDesc(Index).Text
        Case 1

            '   frmMapProductInfo.txtInfo.Text = txtMapDesc(Index).Text
        Case 2

            ' frmMapProductInfo.txtInfo.Text = txtMapDesc(Index).Text
        Case 3
            ' frmMapProductInfo.txtInfo.Text = txtMapDesc(Index).Text
    End Select
    
    frmMapProductInfo.Show vbModeless, Me
    
    ' cmdInfo(Index).SetFocus
    
End Sub

Private Sub cmdInsertIncident_Click()
        
    RaiseEvent InsertIncident
End Sub

Private Sub cmdIOMJOC_Click()
    RaiseEvent ShowIOM
End Sub

Private Function FileExists(Filename As String) As Integer
    Dim i As Integer
    
    On Local Error Resume Next
    i = Len(Dir$(Filename$))

    If Err Or i = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If

    On Local Error GoTo 0
End Function

Private Sub cmdOpenMap_Click(Index As Integer)
    '        '<EhHeader>
    '        On Error GoTo cmdOpenMap_Click_Err
    '        '</EhHeader>
    '        Dim sMaps() As String
    '        Dim sFilePath As String
    '        Dim oExtent As New XGIS_Extent
    '
    '        If g_RSMaps.State = adStateOpen Then
    '100         If Not g_RSAppSettings.EOF Or Not g_RSAppSettings.Bof Then
    '
    '102             SafeMoveFirst g_RSMaps
    '104         '    g_RSMaps.Find "Alias = '" & FraMaps(Index).caption & "'"
    '
    '106             If Not g_RSMaps.EOF Then
    '
    '                    sFilePath = g_sAppPath & "\data\user\Maps\" & g_RSMaps.Fields.Item("FileName").Value
    '
    '108                 If FileExists(sFilePath) Then
    '110                     RaiseEvent LoadMap(g_RSMaps.Fields.Item("FileName").Value, oExtent)
    '                    Else
    '112                     MsgBox "Map filename '" & sFilePath & "' does not exist"
    '
    '                    End If
    '
    '                Else
    '114              '   MsgBox "Alias = '" & FraMaps(Index).caption & "' was not found"
    '
    '                End If
    '            End If
    '        End If
    '
    '        '<EhFooter>
    '        Exit Sub
    '
    'cmdOpenMap_Click_Err:
    '        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.cmdOpenMap_Click " & "at line " & Erl
    '        Resume Next
    '        '</EhFooter>
End Sub

Private Sub cmdPreviewMap_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo cmdPreviewMap_Click_Err
        '</EhHeader>

        Dim sMaps() As String
        Dim sFilePath As String

        If g_RSMaps.State = adStateOpen Then
100         SafeMoveFirst g_RSMaps
102         '   g_RSMaps.Find "Alias = '" & FraMaps(Index).caption & "'"
        
104         If Not g_RSMaps.EOF Then
        
106             sFilePath = g_sAppPath & "\data\user\Maps\Preview\" & g_RSMaps.Fields.Item("Image").value

108             If FileExists(sFilePath) Then

110                 frmMapproductPreview.Image1.Picture = LoadPicture(sFilePath)
112                 frmMapproductPreview.Show vbModal, Me
        
                Else
114                 MsgBox "Map preview '" & sFilePath & "' does not exist"
        
                End If

            Else
        
116             '       MsgBox "Alias = '" & FraMaps(Index).caption & "' was not found"
        
            End If
        End If

        '<EhFooter>
        Exit Sub

cmdPreviewMap_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.cmdPreviewMap_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdReset_Click()
    RaiseEvent ResetSecurityAnalysis
End Sub

Private Sub cmdSecurityAnalysis_Click()
    RaiseEvent DoSecurityAnalysis
End Sub

Private Sub cmdSetScope_Click()
    RaiseEvent setScope
End Sub

Private Sub cmdSetScoring_Click()
    RaiseEvent ScoringSettings
End Sub

Private Sub comGovernorates_Click()
    '    'SELECT CommodityName AS Type,  Delivered FROM Commodity WHERE Admin2 = '" & comGovernorates.list(comGovernorates.listindex) & "'"
    '    dxTypeScoring.Dataset.Close
    '
    '    dxTypeScoring.Dataset.ADODataset.CommandText = "SELECT CommodityName AS Type, Delivered, Need FROM Commodity WHERE Admin2 = '" & comGovernorates.List(comGovernorates.ListIndex) & "'"
    '    dxTypeScoring.Dataset.Open

End Sub

Private Sub ComThematics_Click()
        '<EhHeader>
        On Error GoTo ComThematics_Click_Err
        '</EhHeader>
100     elLegend.Grid(gsRowHeight, 1) = (765)
    
102     If ComThematics.ListIndex = 0 Then
104         RaiseEvent DeActivateTheme(m_sTheme, m_sThemeGR)
        Else
106         m_sTheme = ComThematics.List(ComThematics.ListIndex)
108         m_sThemeGR = ComThematicsGroup.List(ComThematicsGroup.ListIndex)
110         RaiseEvent ActivateTheme(m_sTheme, m_sThemeGR)
        End If
    
        '<EhFooter>
        Exit Sub

ComThematics_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.ComThematics_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComThematicsGroup_Click()
        '<EhHeader>
        On Error GoTo ComThematicsGroup_Click_Err
        '</EhHeader>
        Dim bHasThemes As Boolean
        
100     elLegend.Grid(gsRowHeight, 1) = (765)
    
102     ComThematics.Clear
104     ComThematics.AddItem "--- No Thematics ---"
    
106     If ComThematicsGroup.ListIndex = 0 Then Exit Sub
            
        Dim rsTHGR As New ADODB.Recordset
        
120     rsTHGR.Open "SELECT * FROM ThemeGroups WHERE Name = '" & ComThematicsGroup.List(ComThematicsGroup.ListIndex) & "'", m_Cnn, adOpenDynamic, adLockReadOnly
        
122     Set g_RSThemeSettings = New ADODB.Recordset
124     g_RSThemeSettings.Open "SELECT * FROM Themes WHERE ThemeGroup = " & rsTHGR.Fields.Item("ID").value, m_Cnn, adOpenDynamic, adLockReadOnly

126     If Not g_RSThemeSettings.EOF And Not g_RSThemeSettings.Bof Then
128         bHasThemes = True
130         ComThematics.Clear
132         ComThematics.AddItem "--- Choose Theme ---"
134         SafeMoveFirst g_RSThemeSettings
                    
136         Do While Not g_RSThemeSettings.EOF
138             ComThematics.AddItem g_RSThemeSettings.Fields.Item("Name").value
140             g_RSThemeSettings.MoveNext
            Loop
                    
        End If
         
142     If ComThematics.ListCount > 0 Then ComThematics.ListIndex = 0
  
        '<EhFooter>
        Exit Sub

ComThematicsGroup_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.ComThematicsGroup_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDBGrid1_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, _
                                   ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    
    If C1TabOpsView.CurrTab = 3 Then
        If Not Node.HasChildren Then
            RaiseEvent ZoomToLoc(Node.values(2), Node.values(0))
            'MsgBox "Zoom In To Community:" & Node.Values(2) & " ID:" & Node.Values(0)
        Else
            DebugPrint Node.values(0)
            
        End If
    End If

End Sub

Private Sub dxDBGrid1_OnSelectedCountChange()
        '<EhHeader>
        On Error GoTo dxDBGrid1_OnSelectedCountChange_Err
        '</EhHeader>
        Dim i As Integer
        Dim Node As dxGridNode
        Dim bookmark As Variant
        Dim sVals As String

        'DebugPrint "SEL COUNT CHANGE " & Now & dxDBGrid1.ex.SelectedCount
    
100     If Not dxDBGrid1.M.IsGridMode Then

102         For i = 0 To dxDBGrid1.ex.SelectedCount - 1
104             Set Node = dxDBGrid1.ex.SelectedNodes(i)

106             If sVals = "" Then
108                 sVals = Node.values(0) ' SelectedNodes(i)
                Else
110                 sVals = sVals & "," & Node.values(0)
                End If

            Next

112         ZoomTofeature sVals
        Else

114         For i = 0 To dxDBGrid1.ex.SelectedCount - 1
116             bookmark = dxDBGrid1.ex.SelectedRows(i)

118             If dxDBGrid1.Dataset.BookmarkValid(bookmark) Then
                    'DebugPrint dxDBGrid1.Dataset.FieldValues(dxDBGrid1.Dataset.FieldNameByNo(0))
120                 dxDBGrid1.Dataset.GotoBookmark bookmark
122                 DebugPrint dxDBGrid1.Dataset.FieldValues(dxDBGrid1.Dataset.FieldNameByNo(0))
                
                    '...
                End If

                'DebugPrint node.Values(0) ' SelectedNodes(i)
            Next

        End If

        'dxDBGrid1.m.CopySelectedToClipboard

        '<EhFooter>
        Exit Sub

dxDBGrid1_OnSelectedCountChange_Err:
        MsgBox Err.Description & vbCrLf & "in OASISMAScoring.frmScoring.dxDBGrid1_OnSelectedCountChange " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hdc As Long, _
                                       ByVal left As Single, _
                                       ByVal top As Single, _
                                       ByVal right As Single, _
                                       ByVal bottom As Single, _
                                       ByVal Node As DXDBGRIDLibCtl.IdxGridNode, _
                                       ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, _
                                       ByVal Selected As Boolean, _
                                       ByVal Focused As Boolean, _
                                       ByVal NewItemRow As Boolean, _
                                       Text As String, _
                                       color As Long, _
                                       ByVal Font As stdole.IFontDisp, _
                                       FontColor As Long, _
                                       Alignment As DXDBGRIDLibCtl.ExAlignment, _
                                       Done As Boolean)
        '<EhHeader>
        On Error GoTo dxDBGrid1_OnCustomDrawCell_Err
        '</EhHeader>
        Dim s As String
        Dim q As Integer
        Dim iNumLevel1 As Integer
 
        '       If Node.HasChildren Then 'Kollar om de ar gruperade...
        '
        '            If Node.Level = 1 Then
        '                For q = 0 To Node.Count - 1
        '                    iNumLevel1 = iNumLevel1 + Node.Items(q).Count
        '                    DebugPrint Node.Strings(0) & Node.Items(q).Count
        '
        '                Next q
        '                DebugPrint Node.Strings(0) & " " & iNumLevel1 & " Impacted Communities"
        '            End If
        '
        '
        '        End If
 
        On Error Resume Next
 
100     s = Node.values(dxDBGrid1.Columns.ColumnByFieldName("Impact").Index)
  
102     Select Case s

            Case "High"
                'Color = vbRed
104             color = LColor1(0)

106         Case "Medium"
                'Color = vbBlue
108             color = LColor1(3)

110         Case "Low"
                'Color = vbCyan
112             color = LColor1(1)

114         Case "None"
                'Color = vbGreen
116             color = LColor1(2)
        End Select

        Exit Sub

dxDBGrid1_OnCustomDrawCell_Err:
        MsgBox Err.Description & vbCrLf & "in OASISMAScoring.frmScoring.dxDBGrid1_OnCustomDrawCell " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub DxHAnalyze_Click()
    tabMAOptions.CurrTab = 0
    tabMAOptions.TabVisible(0) = True
End Sub

Private Sub DxHSettings_Click()
    tabMAOptions.CurrTab = 2
End Sub

Private Sub DxHSynchronize_Click()
    MsgBox "Your Server Settings are Invalid! Contact Your OASIS system Administrator for Assistance: " & vbCrLf & " Invalid Server Synch String: http://localhost/DanielSynchTest/ValidateUser?id=js62'353¨+52uw24uthgfdgdfgf%&#%%&", vbInformation, "UNDP Mine Action Data Synch"
End Sub

Private Sub DxHUpdate_Click()
        '<EhHeader>
        On Error GoTo DxHUpdate_Click_Err
        '</EhHeader>
        Dim i As Integer
        Dim bSelected As Boolean
        Dim sURL As String
        Dim iCurRegion As Integer

100     For i = 0 To OptRegion1.UBound

102         If OptRegion1(i).value Then
104             iCurRegion = i + 1
                Exit For
            End If

        Next
    
106     For i = 1 To lstUpdates.ListItems.Count

108         If lstUpdates.ListItems(i).Checked Then
110             bSelected = True
                Exit For
            End If

        Next

112     If Not bSelected Then
114         MsgBox "No data selected in the list below to update data."
        Else
        
            'sURL = m_MASynchURLs(0) & m_MASynchURLs(1) & "?UserName=John&Passwd=john"
            'DebugPrint RunWebFunction(sURL)
            Dim MaFiles() As String
            Dim MaFileAliases() As String
        
116         SafeMoveFirst g_RSAppSettings
118         g_RSAppSettings.Find "SettingName = 'MADataSetFiles'"
120         MaFiles = Split(g_RSAppSettings.Fields.Item("SettingValue1").value, ",")
122         MaFileAliases = Split(g_RSAppSettings.Fields.Item("SettingValue2").value, ",")
        
124         frmDownLoader.Show vbModeless, Me
        
126         frmDownLoader.ClearPrevs
        
128         For i = 1 To lstUpdates.ListItems.Count
    
130             If lstUpdates.ListItems(i).Checked Then
                    On Error Resume Next
                
132                 Kill g_sAppPath & "\data\temp\" & MaFiles(i - 1) & ".shp"
134                 Kill g_sAppPath & "\data\temp\" & MaFiles(i - 1) & ".dbf"
136                 Kill g_sAppPath & "\data\temp\" & MaFiles(i - 1) & ".fld"
138                 Kill g_sAppPath & "\data\temp\" & MaFiles(i - 1) & ".sbn"
140                 Kill g_sAppPath & "\data\temp\" & MaFiles(i - 1) & ".sbx"
                
142                 DebugPrint m_MASynchURLs(0) & "SynchFiles/import/region2/" & MaFiles(i - 1) & ".shp"
                
144                 frmDownLoader.Download_File m_MASynchURLs(0) & "SynchFiles/import/Region" & iCurRegion & "/" & MaFiles(i - 1) & ".shp", g_sAppPath & "\data\temp\Region" & iCurRegion & "\" & MaFiles(i - 1) & ".shp", MaFileAliases(i - 1), g_sAppPath & "\data\gis\vector\Mine_Action\_shp_imsmaoasisshapeexports\" & MaFiles(i - 1) & ".shp"
146                 frmDownLoader.Download_File m_MASynchURLs(0) & "SynchFiles/import/Region" & iCurRegion & "/" & MaFiles(i - 1) & ".dbf", g_sAppPath & "\data\temp\Region" & iCurRegion & "\" & MaFiles(i - 1) & ".dbf", MaFileAliases(i - 1), g_sAppPath & "\data\gis\vector\Mine_Action\_shp_imsmaoasisshapeexports\" & MaFiles(i - 1) & ".dbf"
148                 frmDownLoader.Download_File m_MASynchURLs(0) & "SynchFiles/import/Region" & iCurRegion & "/" & MaFiles(i - 1) & ".fld", g_sAppPath & "\data\temp\Region" & iCurRegion & "\" & MaFiles(i - 1) & ".fld", MaFileAliases(i - 1), g_sAppPath & "\data\gis\vector\Mine_Action\_shp_imsmaoasisshapeexports\" & MaFiles(i - 1) & ".fld"
150                 frmDownLoader.Download_File m_MASynchURLs(0) & "SynchFiles/import/Region" & iCurRegion & "/" & MaFiles(i - 1) & ".sbn", g_sAppPath & "\data\temp\Region" & iCurRegion & "\" & MaFiles(i - 1) & ".sbn", MaFileAliases(i - 1), g_sAppPath & "\data\gis\vector\Mine_Action\_shp_imsmaoasisshapeexports\" & MaFiles(i - 1) & ".sbn"
152                 frmDownLoader.Download_File m_MASynchURLs(0) & "SynchFiles/import/Region" & iCurRegion & "/" & MaFiles(i - 1) & ".sbx", g_sAppPath & "\data\temp\Region" & iCurRegion & "\" & MaFiles(i - 1) & ".sbx", MaFileAliases(i - 1), g_sAppPath & "\data\gis\vector\Mine_Action\_shp_imsmaoasisshapeexports\" & MaFiles(i - 1) & ".sbx"
            
                End If
    
            Next
        
            'MsgBox "Your Server Settings are Invalid! Contact Your OASIS system Administrator for Assistance: " & vbCrLf & " Invalid Server Synch String: http://localhost/DanielSynchTest/ValidateUser?id=js62'353¨+52uw24uthgfdgdfgf%&#%%&", vbInformation, "UNDP Mine Action Data Synch"
        End If

        '<EhFooter>
        Exit Sub

DxHUpdate_Click_Err:
        frmLog.txtLog.Text = Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.DxHUpdate_Click " & "at line " & Erl & vbCrLf & frmLog.txtLog.Text
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxTypeScoring_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, _
                                       ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

    If C1TabOpsView.CurrTab = 7 Then
        If Not Node.HasChildren Then
            RaiseEvent UpdateNFI(comGovernorates.List(comGovernorates.ListIndex), Node.values(1), Node.values(2))
           
        Else
            DebugPrint Node.values(0)
            
        End If
    End If

End Sub

'Private Sub Inet1_StateChanged(ByVal State As Integer)
'Dim DisplayStatus As String
'
'    Select Case State
'        Case icConnected
'            DisplayStatus = "Connected"
'
'        Case icConnecting
'            DisplayStatus = "Connecting"
'
'        Case icDisconnected
'            DisplayStatus = "Disconnected"
'
'        Case icDisconnecting
'            DisplayStatus = "Disconnecting"
'
'        Case icError
'            DisplayStatus = "Error: " & Inet1.ResponseInfo
'
'        Case icReceivingResponse
'            DisplayStatus = "Receiving response"
'
'        Case icRequesting
'            DisplayStatus = "Requesting"
'
'        Case icRequestSent
'            DisplayStatus = "Request Sent"
'
'        Case icResolvingHost
'            DisplayStatus = "Resolving host"
'
'        Case icResponseCompleted
'            DisplayStatus = "Response completed"
'
'        Case icResponseReceived
'            DisplayStatus = "Response received"
'    End Select
'
'    lblIMSMADataUpdate.caption = DisplayStatus
'
'End Sub

Public Function RunWebFunction(ByVal URL As String) As String
    On Error Resume Next
    
    '    MsXmlHttp.Open "GET", URL, 0
    '    MsXmlHttp.Send
    '    RunWebFunction = MsXmlHttp.responseText

    '
    '  On Error Resume Next
    '    Inet1.URL = URL
    '    DebugPrint Inet1.URL
    '    RunWebFunction = Inet1.OpenURL
End Function

Private Sub DxHUpdateData_Click()
    tabMAOptions.CurrTab = 1
End Sub

Private Sub Form_Load()
    PrepareTV
    RaiseEvent UpdateValues
End Sub

Public Sub Init(oGIS As TatukGIS_XDK10.XGIS_Viewer, _
                CN As ADODB.Connection)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        
100     Set mOCN = CN
102     Legend1.GIS_Viewer = oGIS
    
104     LColor1 = Array(RGB(255, 200, 200), RGB(200, 100, 200), RGB(200, 255, 200), RGB(200, 200, 255))
106     lColor = Array(vbRed, vbWhite, vbGreen, vbBlue)
            
108     SafeMoveFirst g_RSAppSettings
110     g_RSAppSettings.Find "SettingName = 'ShowLISTab'"
112     C1TabOpsView.TabVisible(3) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = "1", True, False)

114     SafeMoveFirst g_RSAppSettings
116     g_RSAppSettings.Find "SettingName = 'ShowMapLibraryTab'"

        'override
118     C1TabOpsView.TabVisible(6) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = "1", True, False)

120     If g_RSAppSettings.Fields.Item("SettingValue1").value = "1" Then LoadMaps

122     SafeMoveFirst g_RSAppSettings
126     C1TabOpsView.TabVisible(7) = False 'IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = "1", True, False)

        '128     If g_RSAppSettings.Fields.Item("SettingValue1").Value = "1" Then
        '
        '130         With dxTypeScoring
        '132             .Dataset.Close
        '134             .Dataset.ADODataset.ConnectionString = CN.ConnectionString
        '136             .Dataset.ADODataset.CommandType = cmdText
        '138             .Dataset.ADODataset.CommandText = "SELECT DISTINCT CommodityName AS Type,  Delivered, Need FROM Commodity"
        '140             .Dataset.Open
        '142             .Dataset.Active = True
        '
        '            End With
        '
        '144         'LoadNFI
        '        End If
        '
        '146     If g_RSAppSettings.Fields.Item("SettingValue1").Value = "1" Then
        '
        '148         With dxDBGrid1
        '150             .Event = 1 'EGOnCustomDrawCell
        '152             .EventEnabled = True
        '154             .Options.Set 18 'egoAutoWidth
        '156             .Columns.ApplyBestFit Nothing
        '158             .Dataset.ADODataset.ConnectionString = CN.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\IM.mdb;Persist Security Info=False"
        '
        '160             SafeMoveFirst g_RSAppSettings
        '162             g_RSAppSettings.Find "SettingName = 'MAAddonDataSQL'"
        '
        '164             .Dataset.ADODataset.CommandText = g_RSAppSettings.Fields.Item("SettingValue1").Value '"SELECT ID, Impact, Town, Province, District FROM Scoring ORDER BY Scoring DESC"
        '166             .Dataset.ADODataset.CommandType = cmdText
        '168             .Dataset.Open
        '
        '170             SafeMoveFirst g_RSAppSettings
        '172             g_RSAppSettings.Find "SettingName = 'MAAddonDataKeyField'"
        '
        '174             .KeyField = g_RSAppSettings.Fields.Item("SettingValue1").Value
        '176             .Dataset.Active = True
        '178             .Columns.RetrieveFields
        '180             .Dataset.ADODataset.Requery
        '182             .Columns.Item(0).Visible = False 'ID
        '184             .M.AddGroupColumn .Columns.Item(1)
        '186             .M.AddGroupColumn .Columns.Item(3)
        '188             .M.AddGroupColumn .Columns.Item(4)
        '            End With
        '
        '        End If
    
190     SafeMoveFirst g_RSAppSettings
192     g_RSAppSettings.Find "SettingName = 'ShowMATab'"
194     C1TabOpsView.TabVisible(2) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = "1", True, False)

196     If g_RSAppSettings.Fields.Item("SettingValue1").value = "1" Then
198         SafeMoveFirst g_RSAppSettings
200         g_RSAppSettings.Find "SettingName = 'MineActionThemeMap'"
        
202         GIS.Open g_sAppPath & g_RSAppSettings.Fields.Item("SettingValue1").value, False
204         lgdTest.GIS_Viewer = GIS.viewer
             
            Dim j As Integer
    
206         For j = 0 To GIS.items.Count - 1
208             GIS.items.Item(j).Active = False
            Next
        
210         lgdTest.UpDate
        
212         SafeMoveFirst g_RSAppSettings
214         g_RSAppSettings.Find "SettingName = 'MADataSetFiles'"
            
            Dim sMaDataFiles() As String
            Dim itmAdd As ListItem
            
216         sMaDataFiles = Split(g_RSAppSettings.Fields.Item("SettingValue2").value, ",")
            
218         For j = LBound(sMaDataFiles) To UBound(sMaDataFiles)
220             Set itmAdd = lstUpdates.ListItems.Add(Text:=sMaDataFiles(j))
222             itmAdd.SubItems(1) = "05/07/2007"
224             itmAdd.SubItems(2) = "0000"
            Next
            
226         SafeMoveFirst g_RSAppSettings
228         g_RSAppSettings.Find "SettingName = 'MADataSetFilesUpdate'"
            
230         If Len(g_RSAppSettings.Fields.Item("SettingValue1").value) > 0 Then
232             lblIMSMADataUpdate.caption = "Latest Update:" & vbCrLf & g_RSAppSettings.Fields.Item("SettingValue1").value
            Else
234             lblIMSMADataUpdate.caption = "No updates yet."
            End If
           
236         lstUpdates.ColumnHeaders.Item(1).Width = TextWidth("Mine Action Data Sets Like survey")
238         lstUpdates.ColumnHeaders.Item(2).Width = TextWidth("2007/09/09")
240         lstUpdates.ColumnHeaders.Item(3).Width = TextWidth("0.54532 KB")
            
242         SafeMoveFirst g_RSAppSettings
244         g_RSAppSettings.Find "SettingName = 'MASynchURLs'"
246         m_MASynchURLs(0) = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue1").value) > 0, g_RSAppSettings.Fields.Item("SettingValue1").value, "")
248         m_MASynchURLs(1) = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue2").value) > 0, g_RSAppSettings.Fields.Item("SettingValue2").value, "")
250         m_MASynchURLs(2) = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue3").value) > 0, g_RSAppSettings.Fields.Item("SettingValue3").value, "")
252         m_MASynchURLs(3) = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue4").value) > 0, g_RSAppSettings.Fields.Item("SettingValue4").value, "")
254         m_MASynchURLs(4) = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue5").value) > 0, g_RSAppSettings.Fields.Item("SettingValue5").value, "")
256         m_MASynchURLs(5) = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue6").value) > 0, g_RSAppSettings.Fields.Item("SettingValue6").value, "")
258         m_MASynchURLs(6) = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue7").value) > 0, g_RSAppSettings.Fields.Item("SettingValue7").value, "")
260         m_MASynchURLs(7) = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue8").value) > 0, g_RSAppSettings.Fields.Item("SettingValue8").value, "")
262         m_MASynchURLs(8) = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue9").value) > 0, g_RSAppSettings.Fields.Item("SettingValue9").value, "")
264         m_MASynchURLs(9) = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue10").value) > 0, g_RSAppSettings.Fields.Item("SettingValue10").value, "")
     
266         SafeMoveFirst g_RSAppSettings
268         g_RSAppSettings.Find "SettingName = 'MADefaultRegion'"
270         OptRegion1(CInt(g_RSAppSettings.Fields.Item("SettingValue1").value)).value = True
            
272         SafeMoveFirst g_RSAppSettings
274         g_RSAppSettings.Find "SettingName = 'MAAddonRegions'"
            
            Dim sAddonRegions() As String
            
276         If Not g_RSAppSettings.EOF Then
            
278             sAddonRegions = Split(g_RSAppSettings.Fields.Item("SettingValue1").value, ",")
            
280             For j = LBound(sAddonRegions) To UBound(sAddonRegions)
282                 chkSource(j).caption = sAddonRegions(j)
                Next
            
284             If chkSource.UBound > j Then
    
286                 Do While Not chkSource.UBound = j
288                     chkSource(j).Visible = False
290                     j = j + 1
                    Loop
    
                End If
            End If
        End If
        
        
292     SafeMoveFirst g_RSAppSettings
294     g_RSAppSettings.Find "SettingName = 'ShowSecurityTab'"
296     C1TabOpsView.TabVisible(1) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = "1", True, False)
        
298     If g_RSAppSettings.Fields.Item("SettingValue1").value = "1" Then
            LgdSecurity.Visible = True
            
            With g_RSAppSettings
        
6110            .Requery
6112            SafeMoveFirst g_RSAppSettings
6114            .Find "SettingName = 'AdmProvSec'"
    
6115            If Not .Fields.Item("SettingValue6").value = vbNullString Then OptLevel(0).caption = .Fields.Item("SettingValue6").value
                SafeMoveFirst g_RSAppSettings
                .Find "SettingName = 'AdmDistSec'"

                If Not .Fields.Item("SettingValue6").value = vbNullString Then OptLevel(1).caption = .Fields.Item("SettingValue6").value
                If Not .Fields.Item("SettingValue7").value = vbNullString Then OptLevel(2).caption = .Fields.Item("SettingValue7").value

            End With
            
300         LOADINCDATA
302         CategorizeIncidentsByType
304         abGridTools.Tools.Item("comDateFrom").CBListIndex = 1
        End If
        
        ' LgdSecurity.Options
306     LgdSecurity.GIS_Viewer = GISSecurity.viewer
308     GISSecurity.Mode = TatukGIS_XDK10.XgisZoomEx
        GISSecurity.UpDate
        LgdSecurity.UpDate

310     If Not m_oSQLIncLyr Is Nothing Then m_oSQLIncLyr.Active = False

312     SafeMoveFirst g_RSAppSettings
314     g_RSAppSettings.Find "SettingName = 'ShowlegendTab'"
316     C1TabOpsView.TabVisible(0) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = "1", True, False)
     
318     SafeMoveFirst g_RSAppSettings
320     g_RSAppSettings.Find "SettingName = 'InitialOperationsTabNumber'"

322     If Not bDoNotChangeTab Then C1TabOpsView.CurrTab = CInt(g_RSAppSettings.Fields.Item("SettingValue1").value)

324     SafeMoveFirst g_RSAppSettings
326     g_RSAppSettings.Find "SettingName = 'LgdSecuritySpacing'"
328     LgdSecurity.Spacing = CInt(g_RSAppSettings.Fields.Item("SettingValue1").value)

330     SafeMoveFirst g_RSAppSettings
332     g_RSAppSettings.Find "SettingName = 'LgdMASpacing'"
334     lgdTest.Spacing = CInt(g_RSAppSettings.Fields.Item("SettingValue1").value)
    
        InitSearch oGIS

        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.init " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GetMapsUpdates(strUserGroupPrefix As String)

    Dim rsRemote As ADODB.Recordset
    Dim RS As ADODB.Recordset
    Dim j As Integer
    Dim sString As String
    'Dim RSUpdater As ADODB.Recordset
        
    If Not strUserGroupPrefix = "" Then
                
        Set rsRemote = New ADODB.Recordset
    
        'sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM " & strUserGroupPrefix & "Maps")
        'Set rsRemote = OpenSilentHttpCommsRS(sString, True)
        Set rsRemote = OpenServerRSCompressed(g_sAppServerPath & "/oasis4.asp", "id", "SELECT * FROM " & strUserGroupPrefix & "Maps")
        
        If Not rsRemote Is Nothing Then
            
            If Not rsRemote.State = 0 Then

                m_Cnn.Execute "delete from Maps"
                
                If rsRemote.EOF And rsRemote.Bof Then
                    Exit Sub
                End If
            
                Set RS = New ADODB.Recordset
                SafeMoveFirst rsRemote
                RS.Open "SELECT * FROM Maps", m_Cnn, adOpenDynamic, adLockOptimistic
    
                Do While Not rsRemote.EOF
                    
                    RS.AddNew

                    For j = 1 To rsRemote.Fields.Count - 1

                        If Not IsNull(rsRemote.Fields.Item(j).value) Then
                            'RS.Fields.Item(j).Value = rsRemote.Fields.Item(j).Value
                            RS.Fields.Item(j).value = rsRemote.Fields(RS.Fields.Item(j).Name).value
                        End If

                    Next

                    RS.UpdateBatch adAffectCurrent
                    rsRemote.MoveNext
                    
                Loop
                
                rsRemote.Close
                RS.Close
    
                Set rsRemote = Nothing
                Set RS = Nothing

                '                Set rsRemote = New ADODB.Recordset
                '
                '                sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT SettingValue6 FROM " & strUserGroupPrefix & "AppSettings WHERE SettingName = 'ProfileSettings'")
                '                Set rsRemote = OpenSilentHttpCommsRS(sString, True)
                '
                '                If Not rsRemote.State = 0 Then
                '
                '                    Set RSUpdater = New ADODB.Recordset
                '                    With RSUpdater
                '
                '                        .Open "SELECT * FROM AppSettings", m_Cnn, adOpenDynamic, adLockBatchOptimistic
                '                        .Find "SettingName = 'ProfileSettings'"
                '
                '                        If Not .EOF Then
                '
                '                            If Not IsNull(rsRemote.Fields.Item("SettingValue6").Value) Then
                '                                .Fields("SettingValue6").Value = rsRemote.Fields.Item("SettingValue6").Value
                '                            Else
                '                                .Fields("SettingValue6").Value = 1
                '                            End If
                '
                '                            .UpdateBatch adAffectCurrent
                '                            .Close
                '                        End If
                '
                '                    End With
                '                    Set RSUpdater = Nothing
                '
                '                    rsRemote.Close
                '                    Set rsRemote = Nothing
                '                End If
                
                SynchProfileSettingWithServer "SettingValue6", strUserGroupPrefix, m_Cnn
                
            End If
            
        Else
            
            MsgBox "Something weird happens here. Click OK to initiate breakpoint"
            Stop
            
        End If
                
    End If

End Sub

Private Sub LoadMaps()
'        '<EhHeader>
'        On Error GoTo LoadMaps_Err
'        '</EhHeader>
'        Dim sMaps() As String
'        Dim i As Integer
'
'100     If g_bMapsUpdate Then
'102         GetMapsUpdates g_sRemoteTablePrefix
'104         g_RSMaps.Requery
'        End If
'
'106     If Not g_RSMaps.EOF And Not g_RSMaps.Bof Then
'108         If Not g_RSMaps.Bof Then SafeMoveFirst g_RSMaps
'
'110         i = 0
'
'112         Do While Not g_RSMaps.EOF
'
'114             If FileExists(g_sAppPath & "\data\user\Maps\" & g_RSMaps.Fields.Item("FileName").Value) Then
'116                 '         FraMaps(i).Visible = True
'118                 '    FraMaps(i).caption = g_RSMaps.Fields.Item("Alias").Value
'120                 '  txtMapDesc(i).Text = g_RSMaps.Fields.Item("Description").Value
'                    On Error Resume Next
'122                 '  Set MapIMG(i).Picture = LoadPicture(g_sAppPath & "\data\user\Maps\Preview\" & g_RSMaps.Fields.Item("ThumbNail").Value)
'124                 ''   FraMaps(i).Visible = True
'126                 i = i + 1
'                End If
'
'128             g_RSMaps.MoveNext
'
'            Loop
'
'        End If
'
'        '
'        '        Next
'
'        '<EhFooter>
'        Exit Sub
'
'LoadMaps_Err:
'        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.LoadMaps " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub LOADINCDATA()
        ' LgdSecurity.Color = vbRed
        '  If Not m_oSQLIncLyr Is Nothing Then
        '<EhHeader>
        On Error GoTo LOADINCDATA_Err
        '</EhHeader>
      
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"

104     If Not g_RSAppSettings.EOF Then
        
106         If GISSecurity.get(g_RSAppSettings.Fields.Item("SettingValue1").value) Is Nothing Then
        
108             If g_bIncidentsV2 Then
        
110                 Set m_oSQLIncLyr = New TatukGIS_XDK10.XGIS_LayerSqlAdo

112                 m_oSQLIncLyr.Name = g_RSAppSettings.Fields.Item("SettingValue1").value
114                 m_oSQLIncLyr.SQLParameter("LAYER") = "dd_" & g_sIncidentsV2DDName & "_qryIncidents"
116                 m_oSQLIncLyr.SQLParameter("DIALECT") = g_sGlobalDialect
118                 m_oSQLIncLyr.SQLParameter("ADO") = g_sIncidentsV2ConnectionString
120                 m_oSQLIncLyr.HideFromLegend = False
122                 GISSecurity.Add m_oSQLIncLyr
        
124             ElseIf m_oSQLIncLyr Is Nothing Then
        
126                 Set m_oSQLIncLyr = New TatukGIS_XDK10.XGIS_LayerSqlAdo
                    ' SafeMoveFirst g_RSAppSettings
                    ' g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"
128                 m_oSQLIncLyr.Name = g_RSAppSettings.Fields.Item("SettingValue1").value
130                 m_oSQLIncLyr.SQLParameter("LAYER") = "oincidents"
132                 m_oSQLIncLyr.SQLParameter("DIALECT") = g_sGlobalDialect
                    'm_oSQLIncLyr.SQLParameter("ADO") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\Oasisclient.mdb;" '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\db\Oasisclient.mdb;pwd=<#mypassword#>"
134                 m_oSQLIncLyr.SQLParameter("ADO") = g_sGlobalConnectionString
                    'New OASIS Incidents
136                 m_oSQLIncLyr.HideFromLegend = False
138                 GISSecurity.Add m_oSQLIncLyr

                End If
            
            End If
        
        End If

        '<EhFooter>
        Exit Sub

LOADINCDATA_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMnuOperations.LOADINCDATA " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub CategorizeIncidents()
        '<EhHeader>
        On Error GoTo CategorizeIncidents_Err
        '</EhHeader>
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector
        Dim SymbolList As New TatukGIS_XDK10.XGIS_SymbolList
        Dim RS As New ADODB.Recordset
        Dim sFont As String
        Dim sDBFieldName As String
        Dim sGISFieldName As String
  
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"
104     Set lL = GISSecurity.get(g_RSAppSettings.Fields.Item("SettingValue1").value)
    
106     Set RS = New ADODB.Recordset
                
108     Set RS.ActiveConnection = g_RSAppSettings.ActiveConnection
                
110     RS.CursorType = adOpenDynamic
112     RS.CursorLocation = g_sGlobalCursorLocation
114     RS.LockType = adLockReadOnly
                
116     Select Case abGridTools.Tools.Item("comIncCategory").Text  'shp.GetField("Type")

            Case "--None--"
118             sDBFieldName = "Incident_Type_Name"
120             sGISFieldName = "Type"
122             RS.Open "SELECT * FROM IncTypeCategory ORDER BY [Incident_Type_Name]"

124         Case "Target"
126             sDBFieldName = "Name"
128             sGISFieldName = "Type"
130             RS.Open "SELECT * FROM IncTarget ORDER BY [NAME]"

132         Case "Time"
134             sDBFieldName = "Incident_Time_Name"
136             sGISFieldName = "TIME00"
138             RS.Open "SELECT * FROM IncTimeCategory ORDER BY [Incident_Time_Name]"

140         Case "Type"
142             sDBFieldName = "Incident_Type_Name"
144             sGISFieldName = "Type"
146             RS.Open "SELECT * FROM IncTypeCategory  ORDER BY [Incident_Type_Name]"
        End Select
        
148     Set shp = lL.FindFirst(GISSecurity.viewer.Extent, "", Nothing, "", True)

150     While Not shp Is Nothing
            'DebugPrint " INCIDENT Type:" & shp.GetField("Type")

152         SafeMoveFirst RS
154         RS.Find sDBFieldName & " = " & "'" & shp.GetField(sGISFieldName) & "'"

156         If Not RS.Bof And Not RS.EOF Then
158             sFont = RS.Fields("Font_Name").value
160             sFont = sFont & ":" & RS.Fields("Ascii").value & ":NORMAL"
            End If
                
162         With shp.Params.Marker
                
164             .color = vbWhite
166             .OutlineColor = vbBlue
168             .Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
170             .Size = 640
                
            End With
        
172         shp.draw
        
174         Set shp = lL.FindNext()
        Wend
    
176     lL.Paint
    
        '<EhFooter>
        Exit Sub

CategorizeIncidents_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.CategorizeIncidents " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub CategorizeIncidentsByType()
        '<EhHeader>
        On Error GoTo CategorizeIncidentsByType_Err
        '</EhHeader>
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector
        Dim SymbolList As New TatukGIS_XDK10.XGIS_SymbolList
        Dim RS As New ADODB.Recordset
        Dim sFont As String
        Dim sDBFieldName As String
        Dim sGISFieldName As String
        Dim sTable As String
        Dim oCn As ADODB.Connection
  
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"
104     Set lL = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").value)
    
106     Set RS = New ADODB.Recordset
                
108     If g_bIncidentsV2 Then
            Set oCn = New ADODB.Connection
            oCn.Open g_sIncidentsV2ConnectionString
            Set RS.ActiveConnection = oCn
        Else
            Set RS.ActiveConnection = g_RSAppSettings.ActiveConnection
        End If
                
110     RS.CursorType = adOpenDynamic
112     RS.CursorLocation = g_sGlobalCursorLocation
114     RS.LockType = adLockReadOnly
                
116     Select Case abGridTools.Tools.Item("comIncCategory").Text  'shp.GetField("Type")

            Case "Target"
118             sDBFieldName = IIf(g_bIncidentsV2, "option", "Name")
120             sGISFieldName = IIf(g_bIncidentsV2, "Target", "Target")
                sTable = IIf(g_bIncidentsV2, "dd_" & g_sIncidentsV2DDName & "_ddEventVictim", "IncTarget")
122             RS.Open "SELECT * FROM " & sTable & " ORDER BY [" & sDBFieldName & "]"

124         Case "Time"
126             sDBFieldName = IIf(g_bIncidentsV2, "Incident_Time_Name", "Incident_Time_Name")
128             sGISFieldName = IIf(g_bIncidentsV2, "TIME00", "TIME00")
                sTable = IIf(g_bIncidentsV2, "dd_" & g_sIncidentsV2DDName & "_ddIncTimeCategory", "IncTimeCategory")
130             RS.Open "SELECT * FROM " & sTable & " ORDER BY [" & sDBFieldName & "]"

132         Case "Type"
134             sDBFieldName = IIf(g_bIncidentsV2, "option", "Incident_Type_Name")
136             sGISFieldName = IIf(g_bIncidentsV2, "Type", "Type")
                sTable = IIf(g_bIncidentsV2, "dd_" & g_sIncidentsV2DDName & "_ddEventType", "IncTypeCategory")
138             RS.Open "SELECT * FROM " & sTable & " ORDER BY [" & sDBFieldName & "]"

140         Case Else
142             sDBFieldName = IIf(g_bIncidentsV2, "option", "Incident_Type_Name")
144             sGISFieldName = IIf(g_bIncidentsV2, "Type", "Type")
                sTable = IIf(g_bIncidentsV2, "dd_" & g_sIncidentsV2DDName & "_ddEventType", "IncTypeCategory")
146             RS.Open "SELECT * FROM " & sTable & " ORDER BY [" & sDBFieldName & "]"

        End Select
                
        '    sDBFieldName = "Incident_Type_Name"
        '    sGISFieldName = "Type"
        '    RS.Open "SELECT * FROM IncTypeCategory"
                
148     SafeMoveFirst RS
        
        If m_oSQLIncLyr Is Nothing Then Exit Sub
        
150     m_oSQLIncLyr.ParamsList.Clear

152     Do While Not RS.EOF

154         With RS.Fields
        
156             m_oSQLIncLyr.ParamsList.Add
158             m_oSQLIncLyr.Params.Query = sGISFieldName & " = '" & .Item(sDBFieldName).value & "'"
                'll.Params.AreaColor=RGB(102:102:102)
                'DebugPrint .Item("Incident_Type_Name").value
160             sFont = .Item("Font_Name").value
162             sFont = sFont & ":" & .Item("Ascii").value & ":NORMAL"
            
164             m_oSQLIncLyr.Params.Marker.color = .Item("bgColor").value 'vbWhite
166             m_oSQLIncLyr.Params.Marker.OutlineColor = .Item("color").value 'vbBlue
168             m_oSQLIncLyr.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")

                SafeMoveFirst g_RSAppSettings
                g_RSAppSettings.Find "SettingName = 'ShowSecurityTab'"
                m_oSQLIncLyr.Params.Marker.Size = .Item("size").value
                '     If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue2").Value) Then
                '         If IsNumeric(g_RSAppSettings.Fields.Item("SettingValue2").Value) Then
                '           m_oSQLIncLyr.Params.Marker.Size = CInt(g_RSAppSettings.Fields.Item("SettingValue2").Value)
                '      Else
                '        m_oSQLIncLyr.Params.Marker.Size = 500
                '    End If

                '   Else
                '       m_oSQLIncLyr.Params.Marker.Size = 500
                '   End If
                
172             m_oSQLIncLyr.Params.Marker.ShowLegend = 1
174             m_oSQLIncLyr.Params.Legend = .Item(sDBFieldName).value
176             RS.MoveNext
            End With

        Loop

        Set oCn = Nothing
        'll.UseFileParams
        
        Exit Sub
        '<EhFooter>
        Exit Sub

CategorizeIncidentsByType_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.CategorizeIncidentsByType " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub UpDate()
    Legend1.UpDate
End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    If C1TabOpsView.CurrTab = 0 Or C1TabOpsView.CurrTab = 1 Then cmdExportLegend.Visible = True
    cmdExportLegend.Move Me.Width - (cmdExportLegend.Width * 2), IIf(C1TabOpsView.CurrTab = 0, 360, 1170)

End Sub

Private Sub LgdSecurity_OnLayerActiveChange(translated As Boolean, _
                                            ByVal layer As Object)
    RaiseEvent ActivateSecurityIncidents(layer.Active)
End Sub

Private Sub lgdTest_OnLayerActiveChange(translated As Boolean, _
                                        ByVal layer As Object)
    RaiseEvent LayerActivatedStatus(layer.Name, layer.Active)
    'DebugPrint "MA ALyer:" & Layer.Name & " VISIBLE:" & Layer.Active
End Sub

Private Sub OptRegion1_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo OptRegion1_Click_Err
        '</EhHeader>
        Dim sURL As String
        Dim i As Integer
        Dim vals() As String
        Dim sVal As String
        Dim sCurrentFile As String
        Dim MaFiles() As String
        Dim MaFileAliases() As String
        'Dim i As Integer
        Dim j As Integer
        Dim fileDetails() As String
        Dim currFileNum As Integer
        Dim totFileSize As Long
        Dim lstItem As ListItem
        Dim RSUpdater As ADODB.Recordset

100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'MADataSetFiles'"
104     MaFiles = Split(g_RSAppSettings.Fields.Item("SettingValue1").value, ",")
106     MaFileAliases = Split(g_RSAppSettings.Fields.Item("SettingValue2").value, ",")
    
108     If Index <> 5 Then
110         sURL = m_MASynchURLs(0) & m_MASynchURLs(1) & "?RegionNumber=" & Index + 1
112         sVal = RunWebFunction(sURL)
         
114         If Len(sVal) > 0 Then
                'vals = Split(sVal, "::")
116             vals = Split(sVal, "::")
            
118             lstUpdates.ListItems.Clear
            
120             For i = LBound(MaFiles) To UBound(MaFiles)
122                 totFileSize = 0
124                 currFileNum = 0
                
126                 For j = LBound(vals) To UBound(vals)

128                     If InStr(vals(j), MaFiles(i)) Then
130                         fileDetails = Split(vals(j), "*")
132                         currFileNum = currFileNum + 1
134                         totFileSize = totFileSize + CLng(fileDetails(1))

136                         If currFileNum = 6 Then
                                Exit For
                            End If
                        End If

                    Next
                
138                 If currFileNum = 6 Then
140                     DebugPrint ""
                
142                     Set lstItem = lstUpdates.ListItems.Add(Text:=MaFileAliases(i))
144                     lstItem.SubItems(1) = fileDetails(2)
146                     lstItem.SubItems(2) = totFileSize
                    
                    End If

                Next

                'shp_da_area_1
                'BinarySearchString vals,
            End If

        Else
            'sDebugPrint RunWebFunction(sURL)
        End If
    
        'm_Cnn.Execute "UPDATE " & g_sAppSettingsTable & " SET SettingValue1 = '" & Index & "' WHERE SettingName = 'MADefaultRegion'"
148     Set RSUpdater = New ADODB.Recordset

150     With RSUpdater

152         .Open "SELECT * FROM " & g_sAppSettingsTable, m_Cnn, adOpenDynamic, adLockBatchOptimistic
154         .Find "SettingName = 'MADefaultRegion'"

156         If Not .EOF Then
158             .Fields("SettingValue1").value = Index
                On Error Resume Next
160             .UpdateBatch adAffectCurrent
162             .Close
            End If

        End With

164     Set RSUpdater = Nothing
    
        '<EhFooter>
        Exit Sub

OptRegion1_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.OptRegion1_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub OptW3_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo OptW3_Click_Err
        '</EhHeader>
        Dim sField As String
        Dim sCommandtext As String
        Dim sKeyField As String
        Dim RS As New ADODB.Recordset
        Dim CN As New ADODB.Connection

100     Select Case Index
    
            Case 0
102             sCommandtext = "SELECT DISTINCT clnOrgGUID FROM tblOrgOperationAreas"
104             sKeyField = "clnGUID"
            
106             RS.Open "SELECT clnName FROM tblOrganisation WHERE (SELECT DISTINCT clnOrgGUID FROM tblOrgOperationAreas)", m_Cnn
            
108         Case 1
110             sCommandtext = "SELECT DISTINCT clnlocationGUID FROM tblOrgOperationAreas"
112             sKeyField = "clnGUID"

114         Case 2
                'sCommandtext = "SELECT DISTINCT clnOrgGUID FROM tblOrgOperationAreas"
        End Select
    
116     With dxW3DBGrid
118         .Event = 1 'EGOnCustomDrawCell
120         .EventEnabled = True
122         .Options.Set 18 'egoAutoWidth
124         .Columns.ApplyBestFit Nothing
126         .Dataset.ADODataset.ConnectionString = CN.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\IM.mdb;Persist Security Info=False"
128         .Dataset.ADODataset.CommandText = sCommandtext
130         .Dataset.ADODataset.CommandType = cmdText
132         .Dataset.Open
134         .KeyField = "ID"
136         .Dataset.Active = True
138         .Columns.RetrieveFields
140         .Dataset.ADODataset.Requery
142         .Columns.Item(0).Visible = False 'ID
            '        .m.AddGroupColumn .Columns.Item(1)
            '        .m.AddGroupColumn .Columns.Item(3)
            '        .m.AddGroupColumn .Columns.Item(4)
        End With
    
        '<EhFooter>
        Exit Sub

OptW3_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMnuOperations.OptW3_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub



