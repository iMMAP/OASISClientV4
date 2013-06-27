VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Begin VB.Form frmGISAttrWiz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GIS Layer Attribute Wizard"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7275
   Icon            =   "frmGISAttrWiz.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TatukGIS_DK.XGIS_ViewerWnd GIS 
      Height          =   705
      Left            =   7440
      TabIndex        =   29
      Top             =   4320
      Visible         =   0   'False
      Width           =   345
      BigExtentMargin =   -10
      RestrictedDrag  =   -1  'True
      CachedPaint     =   -1  'True
      IncrementalPaint=   -1  'True
      FullPaint       =   -1  'True
      CodePage        =   0
      OutCodePage     =   0
      CharSet         =   0
      UseRTree        =   0   'False
      PrinterTileSize =   1024
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
      SelectionPattern=   "frmGISAttrWiz.frx":6852
      SelectionTransparency=   100
      SelectionWidth  =   100
      SelectionOutlineOnly=   0   'False
      OldCachedPaint  =   0   'False
      PrinterModeDraft=   0   'False
      PrinterModeForceBitmap=   0   'False
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
      ParentColor     =   0   'False
      ParentCtl3D     =   0   'False
      Object.Visible         =   -1  'True
      Cursor          =   16
      DoubleBuffered  =   0   'False
      ModeMouseButton =   0
      CursorForUserDefined=   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   5370
      TabIndex        =   39
      Top             =   5070
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   285
      Left            =   6330
      TabIndex        =   38
      Top             =   5070
      Width           =   915
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "AdminTools"
      Height          =   285
      Left            =   6300
      TabIndex        =   27
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   285
      Left            =   3450
      TabIndex        =   24
      Top             =   5070
      Width           =   915
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   285
      Left            =   4410
      TabIndex        =   23
      Top             =   5070
      Width           =   915
   End
   Begin C1SizerLibCtl.C1Tab C1TabMain 
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   1710
      Width           =   7275
      _cx             =   12832
      _cy             =   5847
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
      Caption         =   "Tab&1|New Tab|Tab&2|Tab&3|New Tab|New Tab|New Tab|New Tab"
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   2940
         Left            =   8220
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   330
         Width           =   7185
         _cx             =   12674
         _cy             =   5186
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
         Begin VB.ListBox lstLayers 
            Height          =   2085
            Left            =   30
            Style           =   1  'Checkbox
            TabIndex        =   36
            Top             =   330
            Width           =   3525
         End
         Begin VB.CommandButton cmdLoadMapPrj 
            Caption         =   "Open Map Projects"
            Height          =   495
            Left            =   3600
            TabIndex        =   35
            Top             =   330
            Width           =   1035
         End
         Begin VB.Label lblDUDE 
            Caption         =   "Check Layers To Be Configured:"
            Height          =   285
            Left            =   0
            TabIndex        =   34
            Top             =   60
            Width           =   2655
         End
      End
      Begin C1SizerLibCtl.C1Elastic elMapProj 
         Height          =   2940
         Left            =   7920
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   330
         Width           =   7185
         _cx             =   12674
         _cy             =   5186
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
         Begin VB.Frame FraGISAttribute 
            Caption         =   "GIS Attribute Source:"
            Height          =   1335
            Left            =   30
            TabIndex        =   30
            Top             =   30
            Width           =   2535
            Begin VB.OptionButton optGISSource 
               Caption         =   "Incidents"
               Height          =   345
               Index           =   2
               Left            =   180
               TabIndex        =   41
               Top             =   900
               Width           =   1335
            End
            Begin VB.OptionButton optGISSource 
               Caption         =   "Map Project"
               Height          =   345
               Index           =   1
               Left            =   180
               TabIndex        =   32
               Top             =   570
               Width           =   1605
            End
            Begin VB.OptionButton optGISSource 
               Caption         =   "Single File"
               Height          =   345
               Index           =   0
               Left            =   180
               TabIndex        =   31
               Top             =   240
               Value           =   -1  'True
               Width           =   1605
            End
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2940
         Left            =   45
         TabIndex        =   25
         Top             =   330
         Width           =   7185
         Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
            Height          =   3225
            Left            =   0
            OleObjectBlob   =   "frmGISAttrWiz.frx":11164
            TabIndex        =   26
            Top             =   0
            Width           =   7245
         End
      End
      Begin C1SizerLibCtl.C1Elastic el5 
         Height          =   2940
         Left            =   9720
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   330
         Width           =   7185
         _cx             =   12674
         _cy             =   5186
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
         Begin VB.CheckBox chkVisible 
            Caption         =   "Table is visible"
            Height          =   285
            Left            =   4770
            TabIndex        =   21
            Top             =   0
            Width           =   1875
         End
         Begin VB.TextBox txtGISAttr 
            Height          =   315
            Index           =   3
            Left            =   2100
            TabIndex        =   19
            Top             =   0
            Width           =   2505
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Max Allowed Grid Records"
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   20
            Top             =   120
            Width           =   1875
         End
      End
      Begin C1SizerLibCtl.C1Elastic el4 
         Height          =   2940
         Left            =   9420
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   330
         Width           =   7185
         _cx             =   12674
         _cy             =   5186
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
         Begin VB.CheckBox chkDatasetWarning 
            Caption         =   "Activate Dataset Warning level"
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   2805
         End
         Begin VB.TextBox txtGISAttr 
            Height          =   315
            Index           =   2
            Left            =   4200
            TabIndex        =   16
            Top             =   0
            Width           =   2385
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Warning Level"
            Height          =   195
            Index           =   2
            Left            =   3030
            TabIndex        =   17
            Top             =   120
            Width           =   1035
         End
      End
      Begin C1SizerLibCtl.C1Elastic el3 
         Height          =   2940
         Left            =   9120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   330
         Width           =   7185
         _cx             =   12674
         _cy             =   5186
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
         Begin VB.TextBox txtGISAttr 
            Height          =   315
            Index           =   4
            Left            =   1290
            TabIndex        =   10
            Top             =   360
            Width           =   3465
         End
         Begin VB.CheckBox chkAutoRun 
            Caption         =   "Auto Run URLs"
            Height          =   375
            Left            =   1290
            TabIndex        =   13
            Top             =   60
            Width           =   1515
         End
         Begin VB.CheckBox chkIsURL 
            Caption         =   "Is URL Layer"
            Height          =   345
            Left            =   30
            TabIndex        =   12
            Top             =   90
            Width           =   1305
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "URL Field Name"
            Height          =   195
            Index           =   4
            Left            =   30
            TabIndex        =   11
            Top             =   480
            Width           =   1170
         End
      End
      Begin C1SizerLibCtl.C1Elastic el2 
         Height          =   2940
         Left            =   8820
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   330
         Width           =   7185
         _cx             =   12674
         _cy             =   5186
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
         Begin VB.TextBox txtGISAttr 
            Height          =   315
            Index           =   5
            Left            =   1200
            TabIndex        =   22
            Top             =   0
            Width           =   5025
         End
         Begin VB.ListBox lstExcludedFlds 
            Height          =   2085
            Left            =   1200
            Style           =   1  'Checkbox
            TabIndex        =   9
            Top             =   330
            Visible         =   0   'False
            Width           =   5025
         End
         Begin VB.Label lblFieldsAvailable 
            Caption         =   "Fields Available:"
            Height          =   615
            Left            =   0
            TabIndex        =   40
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Excluded Fields:"
            Height          =   195
            Index           =   5
            Left            =   30
            TabIndex        =   8
            Top             =   30
            Width           =   1155
         End
      End
      Begin C1SizerLibCtl.C1Elastic el1 
         Height          =   2940
         Left            =   8520
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   7185
         _cx             =   12674
         _cy             =   5186
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
         Begin VB.TextBox txtGISAttr 
            Height          =   315
            Index           =   1
            Left            =   1110
            TabIndex        =   37
            Top             =   360
            Width           =   3945
         End
         Begin VB.CommandButton cmdOpenGISFile 
            Caption         =   "..."
            Height          =   285
            Left            =   5400
            TabIndex        =   7
            Top             =   0
            Width           =   405
         End
         Begin VB.TextBox txtGISAttr 
            Height          =   315
            Index           =   0
            Left            =   1110
            TabIndex        =   4
            Top             =   0
            Width           =   3945
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Alias"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   6
            Top             =   480
            Width           =   330
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Layer Name"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmGISAttrWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSLocalUserGroups As New ADODB.Recordset
Dim RSLocalGISAttributes As New ADODB.Recordset
Dim m_sIncidentLayerName As String
Dim bEditEntry As VbMsgBoxResult

Private iCurrRecord As Integer

Private GisUtils As New XGIS_Utils
Dim oLyr As XGIS_LayerAbstract
Private Type udtGISAttribute
    sLayerName As String
    sAlias As String
    sExcludedFields() As String
    bURLlyr As Boolean
    sURLFieldName As String
    lWarningLevel As Long
    lMaxLevel As Long
End Type

Private GISAttrs() As udtGISAttribute

Public Sub setUserGroupsRS(ByVal PassedRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo setUserGroupsRS_Err
        '</EhHeader>
100     Set RSLocalUserGroups = PassedRS
    
        '<EhFooter>
        Exit Sub

setUserGroupsRS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.setUserGroupsRS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdBack_Click()
        '<EhHeader>
        On Error GoTo cmdBack_Click_Err
        '</EhHeader>
100     cmdNext.Caption = "Next"
    
102     With Me.C1TabMain

104         If .CurrTab = 1 Then
106             cmdNext.Caption = "Next"
            Else
108             cmdNext.Caption = "Next"
            End If

110         If Not .CurrTab = 0 Then
112             If .CurrTab = 3 Then
114                 .CurrTab = 1
                Else
116                 .CurrTab = .CurrTab - 1
                End If
            End If

        End With

        '<EhFooter>
        Exit Sub

cmdBack_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.cmdBack_Click " & _
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
               "in OASISRemoteAdmin.frmGISAttrWiz.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdLoadMapPrj_Click()
        '<EhHeader>
        On Error GoTo cmdLoadMapPrj_Click_Err
        '</EhHeader>
        Dim i As Long

        Dim c As New cCommonDialog
    
        On Error Resume Next
100     c.DefaultExt = "*.ttkgp"
102     c.DialogTitle = "Open Map Definition File"
104     c.Filter = "Map Definition Files (*.ttkgp;*.prj)|*.ttkgp;*.prj"
106     c.ShowOpen
    
108     If Not c.fileName = "" Then
110         GIS.open c.fileName
112         lstLayers.Clear

114         For i = 0 To GIS.Items.Count
116             ReDim Preserve GISAttrs(i)
118             lstLayers.AddItem GIS.Items.Item(i).Name
            Next
    
        End If

        '<EhFooter>
        Exit Sub

cmdLoadMapPrj_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.cmdLoadMapPrj_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdNext_Click()
        '<EhHeader>
        On Error GoTo cmdNext_Click_Err
        '</EhHeader>

        Dim bVer As Boolean
        
100     cmdBack.Caption = "Back"

102     With Me.C1TabMain
    
104         Select Case .CurrTab
    
                Case 0

106                 bVer = True
108                 bEditEntry = vbNo

110                 If Not RSLocalGISAttributes.EOF Or Not RSLocalGISAttributes.Bof Then
                    
112                     bEditEntry = MsgBox("Do you want to edit the selected entry?  Clicking 'No' will add a new entry", vbYesNoCancel, "Confirm operation")

114                     If bEditEntry = vbYes Then
116                         Call EditExistingEntry
118                     ElseIf bEditEntry = vbNo Then
120                         Call PrepareForNewEntry
                        Else
122                         bVer = False
                        End If

                    Else
124                     Call PrepareForNewEntry
                    End If

126             Case 1, 4, 5
128                 bVer = True
                
130             Case 2
132                 bVer = True
                    '.CurrTab = .CurrTab + 1
                    
134                 If iCurrRecord <= UBound(GISAttrs) Then
136                     txtGISAttr(0).Text = GISAttrs(iCurrRecord).sLayerName
                    End If
                    
138             Case 3
               
140                 If Len(Me.txtGISAttr(0).Text) > 0 And Len(Me.txtGISAttr(1).Text) > 0 Then

                        'If optGISSource(0).Value Then 'Single File
                        '    SetLyrAttr GIS.Get(txtGISAttr(0).Text).Path
142                     If optGISSource(1).Value Then 'MAP PROJECT
144                         SetLyrAttr GIS.Get(GISAttrs(iCurrRecord).sLayerName).Path
146                     ElseIf optGISSource(2).Value Then 'Special Case Incidents

148                         If Len(m_sIncidentLayerName) > 0 Then
150                             SetIncLyrAttr
                            Else
152                             m_sIncidentLayerName = "N/A"
                            End If
                        End If
                        
154                     bVer = True
                    Else
156                     MsgBox "Please enter a valid name and alias!"
                    End If

158             Case 6

160                 If Not IsNumeric(txtGISAttr(2)) Then
162                     bVer = False
164                     MsgBox "Must be a number!"
                    Else
166                     bVer = True
                    End If

168             Case 7
                    
170                 If IsNumeric(Me.txtGISAttr(3)) Then
172                     If optGISSource(0).Value Or optGISSource(2).Value Then

174                         If bEditEntry = vbNo Then RSLocalGISAttributes.AddNew

176                         With RSLocalGISAttributes.fields
                
178                             .Item("name").Value = txtGISAttr(0)
180                             .Item("alias").Value = txtGISAttr(1)
182                             .Item("excludedFlds").Value = txtGISAttr(5)
184                             .Item("MaxRec").Value = txtGISAttr(3)
186                             .Item("URLLayerField").Value = txtGISAttr(4)
188                             .Item("warninglevel").Value = txtGISAttr(2)
                
190                             .Item("autoRunUrls").Value = IIf(Me.chkAutoRun.Value = vbChecked, True, False)
192                             .Item("datasetwarning").Value = IIf(Me.chkDatasetWarning.Value = vbChecked, True, False)
                                'If bEditEntry = vbNo Then .Item("id").Value = Rnd(1000)
194                             .Item("isURLLayer").Value = IIf(Me.chkIsURL.Value = vbChecked, True, False)
196                             .Item("Visible").Value = IIf(Me.chkVisible.Value = vbChecked, True, False)
                
                            End With
                            
                            'RSLocalGISAttributes.UpdateBatch adAffectCurrent
198                         .CurrTab = 0

                        Else

200                         If iCurrRecord <> UBound(GISAttrs) Then

202                             If bEditEntry = vbNo Then RSLocalGISAttributes.AddNew

204                             With RSLocalGISAttributes.fields

206                                 .Item("name").Value = txtGISAttr(0)
208                                 .Item("alias").Value = txtGISAttr(1)
210                                 .Item("excludedFlds").Value = txtGISAttr(5)
212                                 .Item("MaxRec").Value = txtGISAttr(3)
214                                 .Item("URLLayerField").Value = txtGISAttr(4)
216                                 .Item("warninglevel").Value = txtGISAttr(2)
                
218                                 .Item("autoRunUrls").Value = IIf(Me.chkAutoRun.Value = vbChecked, True, False)
220                                 .Item("datasetwarning").Value = IIf(Me.chkDatasetWarning.Value = vbChecked, True, False)
                                    'If bEditEntry = vbNo Then .Item("id").Value = Rnd(1000)
222                                 .Item("isURLLayer").Value = IIf(Me.chkIsURL.Value = vbChecked, True, False)
224                                 .Item("Visible").Value = IIf(Me.chkVisible.Value = vbChecked, True, False)
                
                                End With
                                
                                'RSLocalGISAttributes.UpdateBatch adAffectCurrent
226                             .CurrTab = 3
228                             bVer = True
230                             Call PrepareForNewEntry
                                
232                             iCurrRecord = iCurrRecord + 1
234                             txtGISAttr(0).Text = GISAttrs(iCurrRecord).sLayerName

                            Else
236                             If bEditEntry = vbNo Then RSLocalGISAttributes.AddNew

238                             With RSLocalGISAttributes.fields

240                                 .Item("name").Value = txtGISAttr(0)
242                                 .Item("alias").Value = txtGISAttr(1)
244                                 .Item("excludedFlds").Value = txtGISAttr(5)
246                                 .Item("MaxRec").Value = txtGISAttr(3)
248                                 .Item("URLLayerField").Value = txtGISAttr(4)
250                                 .Item("warninglevel").Value = txtGISAttr(2)
                
252                                 .Item("autoRunUrls").Value = IIf(Me.chkAutoRun.Value = vbChecked, True, False)
254                                 .Item("datasetwarning").Value = IIf(Me.chkDatasetWarning.Value = vbChecked, True, False)
                                    'If bEditEntry = vbNo Then .Item("id").Value = Rnd(1000)
256                                 .Item("isURLLayer").Value = IIf(Me.chkIsURL.Value = vbChecked, True, False)
258                                 .Item("Visible").Value = IIf(Me.chkVisible.Value = vbChecked, True, False)
                
                                End With

                                'RSLocalGISAttributes.UpdateBatch adAffectCurrent
260                             .CurrTab = 0
                            End If
                        End If

                    Else
262                     MsgBox "Please enter a number for the max record!"
                    End If

264                 bVer = False
                
            End Select

266         If bVer Then
                
268             If Not .CurrTab = .NumTabs Then

270                 If .CurrTab = 1 Then
272                     If optGISSource(0).Value = True Or optGISSource(2).Value = True Then
274                         .CurrTab = .CurrTab + 1

276                         If optGISSource(2).Value = True Then
278                             txtGISAttr(0).Text = m_sIncidentLayerName
280                             txtGISAttr(1).Text = m_sIncidentLayerName
                            End If
                        End If
                    End If
                    
282                 .CurrTab = .CurrTab + 1
                End If
                
            End If

        End With

        '<EhFooter>
        Exit Sub

cmdNext_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.cmdNext_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSave_Click()
        '<EhHeader>
        On Error GoTo cmdSave_Click_Err
        '</EhHeader>
        Dim bReturnValue As Boolean
    
100     RSLocalGISAttributes.Filter = adFilterPendingRecords

102     If Not RSLocalGISAttributes.EOF And Not RSLocalGISAttributes.Bof Then
        
104         If MsgBox("Do you wish to save your changes?", vbYesNo, "Confirm Save") = vbYes Then
        
106             bReturnValue = m_frmOASISProgress.SaveHttpCommsRS(RSLocalGISAttributes, WebSite & "Oasis.asp", True)

108             If bReturnValue Then

110                 IncrementProfileSettingVersion WebSite, "SettingValue4", RSLocalUserGroups.fields("Name").Value
112                 MsgBox "Data saved to server"

                Else
114                 MsgBox "Saving to server failed!"
                End If

            End If

        End If

116     Unload Me
        '<EhFooter>
        Exit Sub

cmdSave_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.cmdSave_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub lstLayers_ItemCheck(Item As Integer)
        '<EhHeader>
        On Error GoTo lstLayers_ItemCheck_Err
        '</EhHeader>
        Dim i As Integer
        Dim j As Integer
        Dim oLyr As XGIS_LayerAbstract

100     If lstLayers.SelCount > 0 Then
102         ReDim GISAttrs(lstLayers.SelCount - 1)
        
104         For i = 0 To lstLayers.ListCount
        
106             If lstLayers.Selected(i) Then
108                 GISAttrs(j).sLayerName = lstLayers.List(i)
110                 j = j + 1

112                 If j = lstLayers.SelCount Then Exit For
                End If
    
            Next

114         lstExcludedFlds.Clear
116         lstExcludedFlds.Visible = True
118         txtGISAttr(5).Visible = True
            
            Exit Sub

120         If lstLayers.Selected(Item) Then
                'txtLyrName.Text = lstLayers.List(Item)
        
122             lstExcludedFlds.Clear
124             lstExcludedFlds.Visible = True
            
126             Set oLyr = GIS.Get(lstLayers.List(Item))
                  
128             With oLyr.fields
            
130                 For i = 0 To .Count - 1
132                     lstExcludedFlds.AddItem .Item(i).Name
134                     lstExcludedFlds.Visible = True
                        'txtGISAttr(5).Visible = False
                    Next
            
                End With
        
            Else
                'txtLyrName.Text = ""
            End If
        End If

        '<EhFooter>
        Exit Sub

lstLayers_ItemCheck_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.lstLayers_ItemCheck " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOpenGISFile_Click()
        '<EhHeader>
        On Error GoTo cmdOpenGISFile_Click_Err
        '</EhHeader>
        
        Dim c As New cCommonDialog
        Dim i As Integer
    
        'On Error Resume Next
100     c.DefaultExt = "*.shp"
102     c.DialogTitle = "Open GIS Vector File"
104     c.Filter = "Layer Files (shp)|" & GisUtils.GisSupportedFiles(XgisFileTypeVector, False)
106     c.ShowOpen

108     If c.fileName <> "" Then

110         SetLyrAttr c.fileName
112         Me.txtGISAttr(0) = c.fileName
        
114         txtGISAttr(0) = Right(txtGISAttr(0), Len(txtGISAttr(0)) - InStrRev(txtGISAttr(0), "\"))
116         txtGISAttr(0) = Left(txtGISAttr(0), InStrRev(txtGISAttr(0), ".") - 1)
        End If
         
        '<EhFooter>
        Exit Sub

cmdOpenGISFile_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.cmdOpenGISFile_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetLyrAttr(sLyrFileName As String)
        '<EhHeader>
        On Error GoTo SetLyrAttr_Err
        '</EhHeader>
        Dim i As Integer

100     If InStr(UCase(sLyrFileName), ".SHP") Then
102         Set oLyr = New XGIS_LayerSHP
104     ElseIf InStr(UCase(sLyrFileName), ".KML") Then
106         Set oLyr = New XGIS_LayerKML
108     ElseIf InStr(UCase(sLyrFileName), ".TAB") Then
110         Set oLyr = New XGIS_LayerTAB
112     ElseIf InStr(UCase(sLyrFileName), ".MIF") Then
114         Set oLyr = New XGIS_LayerMIF
116     ElseIf InStr(UCase(sLyrFileName), ".GML") Then
118         Set oLyr = New XGIS_LayerGML
120     ElseIf InStr(UCase(sLyrFileName), ".DXF") Then
122         Set oLyr = New XGIS_LayerDXF
124     ElseIf InStr(UCase(sLyrFileName), ".TTKLS") Then
126         Set oLyr = New XGIS_LayerSqlAdo
        Else
128         Set oLyr = New XGIS_LayerVector
        End If

130     oLyr.Path = sLyrFileName
132     oLyr.open
        'txtGISAttr(0).Text = oLyr.Name
         
134     lstExcludedFlds.Visible = True
136     txtGISAttr(5).Visible = True
138     txtGISAttr(5).Text = ""
140     lstExcludedFlds.Clear
                        
142     With oLyr.fields
            
144         For i = 0 To .Count - 1
146             lstExcludedFlds.AddItem .Item(i).Name
148             lstExcludedFlds.Visible = True
                'txtGISAttr(5).Visible = IIf(optGISSource(0).Value, False, True)
            Next
            
        End With

        '<EhFooter>
        Exit Sub

SetLyrAttr_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.SetLyrAttr " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetIncLyrAttr()
        '<EhHeader>
        On Error GoTo SetIncLyrAttr_Err
        '</EhHeader>
        Dim i As Integer
        Dim RSIncidentsTable As New ADODB.Recordset
        Dim ooConn As New ADODB.Connection

100     With ooConn
102         .CursorLocation = adUseClient
104         .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CreateAppPath & "\data\db\Oasisclient.mdb"
106         .open
        End With
        
108     RSIncidentsTable.open "SELECT * FROM oIncidents_FEA", ooConn
         
110     lstExcludedFlds.Visible = True
112     txtGISAttr(5).Visible = True
114     txtGISAttr(5).Text = ""
116     lstExcludedFlds.Clear
                        
118     With RSIncidentsTable.fields
            
120         For i = 0 To .Count - 1
122             lstExcludedFlds.AddItem .Item(i).Name
124             lstExcludedFlds.Visible = True
                'txtGISAttr(5).Visible = IIf(optGISSource(0).Value, False, True)
            Next
            
        End With
        
126     Set RSIncidentsTable = Nothing
128     ooConn.Close

        '<EhFooter>
        Exit Sub

SetIncLyrAttr_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.SetIncLyrAttr " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTools_Click()
        '<EhHeader>
        On Error GoTo cmdTools_Click_Err
        '</EhHeader>
        Dim m_frmAdminTools As frmAdminTools
100     Set m_frmAdminTools = New frmAdminTools
102     m_frmAdminTools.Show vbModeless, Me
104     Set m_frmAdminTools = Nothing
        '<EhFooter>
        Exit Sub

cmdTools_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.cmdTools_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDBGrid1_OnDblClick()
        '<EhHeader>
        On Error GoTo dxDBGrid1_OnDblClick_Err
        '</EhHeader>

100    ' If DeleteRecordFromRSAndSave(RSLocalGISAttributes, "SettingValue4", RSLocalUserGroups.fields("Name").Value) Then Unload Me
        DeleteRecordFromRSAndSave RSLocalGISAttributes, "SettingValue4", RSLocalUserGroups.fields("Name").Value

        '<EhFooter>
        Exit Sub

dxDBGrid1_OnDblClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.dxDBGrid1_OnDblClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
        Dim rsOName As New ADODB.Recordset
        Dim sString As String
        Dim sSQL As String
        
        C1TabMain.TabHeight = 1
        C1TabMain.CurrTab = 0
        
100     chkVisible.Value = vbChecked

102     DoEvents
104     Me.Picture = g_PictureDialogSmall
106     Set RSLocalGISAttributes = New ADODB.Recordset
        
108     sSQL = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT SettingValue1 FROM " & RSLocalUserGroups!Name & "AppSettings WHERE SettingName = 'OASIS_Incident_Layer_Name'")
110     Set rsOName = m_frmOASISProgress.OpenHttpCommsRS(sSQL, True)
        
112     If Not rsOName.State = adStateClosed Then
114         If Not rsOName.Bof And Not rsOName.EOF Then
116             m_sIncidentLayerName = rsOName!SettingValue1
            End If
        End If

118     sSQL = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & RSLocalUserGroups!Name & "GISGridTableSettings")
120     Set RSLocalGISAttributes = m_frmOASISProgress.OpenHttpCommsRS(sSQL, True)
     
122     If RSLocalGISAttributes.State = adStateClosed Then
124         MsgBox "[" & RSLocalUserGroups!Name & "GISGridTableSettings] Table Does Not Exist!"
            Exit Sub
        End If

        dxDBGrid1.Columns.DestroyColumns
        dxDBGrid1.KeyField = RSLocalGISAttributes.fields(0).Name
126     Set dxDBGrid1.DataSource = RSLocalGISAttributes
        dxDBGrid1.Columns.RetrieveFields
        dxDBGrid1.Columns(0).Visible = False
        dxDBGrid1.Columns(1).Visible = True
        dxDBGrid1.Columns(1).Caption = "Name"
        dxDBGrid1.Columns(2).Visible = True
        dxDBGrid1.Columns(2).Caption = "Alias"
        dxDBGrid1.Columns(3).Visible = False
        dxDBGrid1.Columns(4).Visible = False
        dxDBGrid1.Columns(5).Visible = False
        dxDBGrid1.Columns(6).Visible = False
        dxDBGrid1.Columns(7).Visible = False
        dxDBGrid1.Columns(8).Visible = False
        dxDBGrid1.Columns(9).Visible = False
        dxDBGrid1.Columns(10).Visible = False
        
128     Set rsOName = Nothing
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CheckIfNull(sString As Variant)

    If IsNull(sString) Then
        CheckIfNull = ""
    Else
        CheckIfNull = sString
    End If

End Function

Private Sub EditExistingEntry()
        '<EhHeader>
        On Error GoTo EditExistingEntry_Err
        '</EhHeader>

100     optGISSource(0).Value = True
102     FraGISAttribute.Enabled = False
    
104     With RSLocalGISAttributes.fields
            
106         txtGISAttr(0) = CheckIfNull(.Item("name").Value)
108         txtGISAttr(1) = CheckIfNull(.Item("alias").Value)
110         txtGISAttr(5) = CheckIfNull(.Item("excludedFlds").Value)
112         txtGISAttr(3) = CheckIfNull(.Item("MaxRec").Value)
114         txtGISAttr(4) = CheckIfNull(.Item("URLLayerField").Value)
116         txtGISAttr(2) = CheckIfNull(.Item("warninglevel").Value)
                
118         chkAutoRun.Value = IIf(.Item("autoRunUrls").Value = True, vbChecked, vbUnchecked)
120         chkDatasetWarning.Value = IIf(.Item("datasetwarning").Value = True, vbChecked, vbUnchecked)
122         chkIsURL.Value = IIf(.Item("isURLLayer").Value = True, vbChecked, vbUnchecked)
124         chkVisible.Value = IIf(.Item("Visible").Value = True, vbChecked, vbUnchecked)
126         chkAutoRun.Value = IIf(.Item("autoRunUrls").Value = True, vbChecked, vbUnchecked)
                
        End With

128     If m_sIncidentLayerName = txtGISAttr(0).Text Then optGISSource(2).Value = True
 
        '<EhFooter>
        Exit Sub

EditExistingEntry_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.EditExistingEntry " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub PrepareForNewEntry()
        '<EhHeader>
        On Error GoTo PrepareForNewEntry_Err
        '</EhHeader>

        FraGISAttribute.Enabled = True
100     Me.chkAutoRun.Value = vbUnchecked
102     Me.chkDatasetWarning.Value = vbUnchecked
104     Me.chkIsURL.Value = vbUnchecked
106     Me.chkVisible.Value = vbUnchecked

        '<EhFooter>
        Exit Sub

PrepareForNewEntry_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.PrepareForNewEntry " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

102     Set RSLocalGISAttributes = Nothing
104     Set RSLocalUserGroups = Nothing
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub lstExcludedFlds_ItemCheck(Item As Integer)
        '<EhHeader>
        On Error GoTo lstExcludedFlds_ItemCheck_Err
        '</EhHeader>

        Dim sExcFlds As String
        Dim i As Integer

100     m_frmDebug.DebugPrint lstExcludedFlds.List(Item) & " Selected:" & lstExcludedFlds.Selected(Item)
     
102     For i = 0 To lstExcludedFlds.ListCount - 1

104         If lstExcludedFlds.Selected(i) Then
106             sExcFlds = sExcFlds & IIf(Len(sExcFlds) > 1, ",", "") & lstExcludedFlds.List(i)
            End If

        Next
    
108     txtGISAttr(5).Text = sExcFlds
    
110     m_frmDebug.DebugPrint sExcFlds
    
        '<EhFooter>
        Exit Sub

lstExcludedFlds_ItemCheck_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.lstExcludedFlds_ItemCheck " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub optGISSource_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo optGISSource_Click_Err
        '</EhHeader>

100     If Index = 0 Then
102         cmdOpenGISFile.Enabled = True
104         txtGISAttr(0).Locked = True
106         txtGISAttr(0).Enabled = True
        Else
108         txtGISAttr(0).Locked = False
110         txtGISAttr(0).Enabled = False
112         cmdOpenGISFile.Enabled = False
114         lstExcludedFlds.Visible = True
116         txtGISAttr(5).Visible = True
        End If

        '<EhFooter>
        Exit Sub

optGISSource_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmGISAttrWiz.optGISSource_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
