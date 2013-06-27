VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Begin VB.Form frmThemeWiz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Theme Wizard"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7275
   Icon            =   "frmThemeWiz.frx":0000
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
      TabIndex        =   20
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
      SelectionPattern=   "frmThemeWiz.frx":6852
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
      TabIndex        =   29
      Top             =   5070
      Width           =   915
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   285
      Left            =   6330
      TabIndex        =   28
      Top             =   5070
      Width           =   915
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "AdminTools"
      Height          =   285
      Left            =   6300
      TabIndex        =   18
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   285
      Left            =   3450
      TabIndex        =   15
      Top             =   5070
      Width           =   915
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   285
      Left            =   4410
      TabIndex        =   14
      Top             =   5070
      Width           =   915
   End
   Begin C1SizerLibCtl.C1Tab C1TabMain 
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   1740
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
         TabIndex        =   24
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
            Height          =   2400
            Left            =   30
            TabIndex        =   27
            Top             =   330
            Width           =   3525
         End
         Begin VB.CommandButton cmdLoadMapPrj 
            Caption         =   "Open Map Projects"
            Height          =   495
            Left            =   3600
            TabIndex        =   26
            Top             =   330
            Width           =   1035
         End
         Begin VB.Label lblDUDE 
            Caption         =   "Select the Analysis Layer:"
            Height          =   285
            Left            =   0
            TabIndex        =   25
            Top             =   60
            Width           =   2655
         End
      End
      Begin C1SizerLibCtl.C1Elastic elMapProj 
         Height          =   2940
         Left            =   7920
         TabIndex        =   19
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
            Caption         =   "Analysis Field Source:"
            Height          =   1335
            Left            =   150
            TabIndex        =   21
            Top             =   180
            Width           =   2535
            Begin VB.OptionButton optGISSource 
               Caption         =   "Incidents"
               Height          =   345
               Index           =   2
               Left            =   180
               TabIndex        =   31
               Top             =   900
               Width           =   1335
            End
            Begin VB.OptionButton optGISSource 
               Caption         =   "Map Project"
               Height          =   345
               Index           =   1
               Left            =   180
               TabIndex        =   23
               Top             =   570
               Width           =   1605
            End
            Begin VB.OptionButton optGISSource 
               Caption         =   "Single File"
               Height          =   345
               Index           =   0
               Left            =   180
               TabIndex        =   22
               Top             =   240
               Value           =   -1  'True
               Width           =   1605
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2940
         Left            =   45
         TabIndex        =   16
         Top             =   330
         Width           =   7185
         Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
            Height          =   3225
            Left            =   0
            OleObjectBlob   =   "frmThemeWiz.frx":11164
            TabIndex        =   17
            Top             =   0
            Width           =   7245
         End
      End
      Begin C1SizerLibCtl.C1Elastic el5 
         Height          =   2940
         Left            =   9720
         TabIndex        =   11
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
         Begin VB.FileListBox txtIniFileName 
            Height          =   2040
            Left            =   240
            Pattern         =   "*.ini"
            TabIndex        =   38
            Top             =   480
            Width           =   6645
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Theme Configuration FileName"
            Height          =   195
            Index           =   3
            Left            =   210
            TabIndex        =   13
            Top             =   240
            Width           =   2175
         End
      End
      Begin C1SizerLibCtl.C1Elastic el4 
         Height          =   2940
         Left            =   9420
         TabIndex        =   10
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
         Begin XpressEditorsLibCtl.dxMemoEdit txtMaps 
            Height          =   1335
            Left            =   390
            OleObjectBlob   =   "frmThemeWiz.frx":11E0C
            TabIndex        =   36
            Top             =   600
            Width           =   6345
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Maps (currently unavailable)"
            Height          =   195
            Index           =   2
            Left            =   390
            TabIndex        =   12
            Top             =   330
            Width           =   1980
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
         Begin XpressEditorsLibCtl.dxMemoEdit txtDescription 
            Height          =   885
            Left            =   1740
            OleObjectBlob   =   "frmThemeWiz.frx":11F20
            TabIndex        =   37
            Top             =   840
            Width           =   3495
         End
         Begin VB.CommandButton cmdAddThemeGroup 
            Caption         =   "..."
            Height          =   285
            Left            =   5310
            TabIndex        =   35
            Top             =   1860
            Width           =   345
         End
         Begin VB.ComboBox cmbThemeGroup 
            Height          =   315
            Left            =   1740
            TabIndex        =   34
            Top             =   1860
            Width           =   3495
         End
         Begin VB.TextBox txtThemeName 
            Height          =   315
            Left            =   1740
            TabIndex        =   8
            Top             =   390
            Width           =   3465
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Theme Group"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   33
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Theme Description"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   32
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Theme Name"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   9
            Top             =   450
            Width           =   960
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
         Begin VB.ListBox lstAnalysisFlds 
            Height          =   2595
            Left            =   1200
            TabIndex        =   7
            Top             =   180
            Visible         =   0   'False
            Width           =   5025
         End
         Begin VB.Label lblFieldsAvailable 
            Caption         =   "Select the Analysis Field:"
            Height          =   615
            Left            =   90
            TabIndex        =   30
            Top             =   210
            Width           =   1095
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
         Begin VB.CommandButton cmdOpenGISFile 
            Caption         =   "..."
            Height          =   285
            Left            =   4110
            TabIndex        =   6
            Top             =   540
            Width           =   405
         End
         Begin VB.TextBox txtLayerName 
            Height          =   315
            Left            =   90
            TabIndex        =   4
            Top             =   540
            Width           =   3945
         End
         Begin VB.Label lblCaption 
            AutoSize        =   -1  'True
            Caption         =   "Analysis Layer Name"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   5
            Top             =   300
            Width           =   1470
         End
      End
   End
End
Attribute VB_Name = "frmThemeWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSLocalUserGroups As New ADODB.Recordset
Dim RSLocalThemes As New ADODB.Recordset
Dim RSLocalThemeGroups As New ADODB.Recordset
Dim m_sIncidentLayerName As String
Dim WithEvents m_frmThemeGroups As frmThemeGroups
Attribute m_frmThemeGroups.VB_VarHelpID = -1
Dim bEditEntry As VbMsgBoxResult

Private iCurrRecord As Integer

Private GisUtils As New XGIS_Utils
Dim oLyr As XGIS_LayerAbstract

Public Sub setUserGroupsRS(ByVal PassedRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo setUserGroupsRS_Err
        '</EhHeader>
100     Set RSLocalUserGroups = PassedRS
        '<EhFooter>
        Exit Sub

setUserGroupsRS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmThemeWiz.setUserGroupsRS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAddThemeGroup_Click()
        '<EhHeader>
        On Error GoTo cmdAddThemeGroup_Click_Err
        '</EhHeader>

100     Set m_frmThemeGroups = New frmThemeGroups
102     m_frmThemeGroups.setUserGroupsRS RSLocalUserGroups
104     m_frmThemeGroups.Show vbModeless, Me
    
        '<EhFooter>
        Exit Sub

cmdAddThemeGroup_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmThemeWiz.cmdAddThemeGroup_Click " & _
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
106             cmdNext.Caption = "New"
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
               "in OASISRemoteAdmin.frmThemeWiz.cmdBack_Click " & _
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
               "in OASISRemoteAdmin.frmThemeWiz.cmdCancel_Click " & _
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
106     c.InitDir = CreateAppPath & "\data\user\maps"
108     c.ShowOpen
    
110     If Not c.fileName = "" Then
112         GIS.open c.fileName
114         lstLayers.Clear

116         For i = 0 To GIS.Items.Count
118             lstLayers.AddItem GIS.Items.Item(i).Name
            Next
    
        End If

        '<EhFooter>
        Exit Sub

cmdLoadMapPrj_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmThemeWiz.cmdLoadMapPrj_Click " & _
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
    
                Case 0, 1

106                 bVer = True

                    bEditEntry = vbNo

                    If Not RSLocalThemes.EOF Or Not RSLocalThemes.Bof Then
                    
                        bEditEntry = MsgBox("Do you want to edit the selected entry?  Clicking 'No' will add a new entry", vbYesNoCancel, "Confirm operation")

                        If bEditEntry = vbYes Then
                            Call EditExistingEntry
                            .CurrTab = .CurrTab + 1
                        ElseIf bEditEntry = vbCancel Then
                            bVer = False
                        End If

                    End If
                
108             Case 2

110                 If Len(lstLayers.Text) > 0 Then
112                     bVer = True
114                     txtLayerName.Text = lstLayers.Text
                    Else
116                     bVer = False
118                     MsgBox "Please select a layer"
                    End If

120             Case 3
               
122                 If Len(Me.txtLayerName.Text) > 0 Then

124                     If optGISSource(1).Value Then 'MAP PROJECT or VectorFile
                    
126                         SetLyrAttr GIS.Get(lstLayers.Text).Path
                        
128                     ElseIf optGISSource(2).Value Then 'Special Case Incidents

130                         If Len(m_sIncidentLayerName) > 0 Then
132                             SetIncLyrAttr
                            Else
134                             m_sIncidentLayerName = "N/A"
                            End If
                        End If
                        
136                     bVer = True
                    Else
138                     MsgBox "Please enter an Analysis Layer name"
                    End If

140             Case 4

142                 If Len(lstAnalysisFlds) > 0 Then
144                     bVer = True
                    Else
146                     bVer = False
148                     MsgBox "Please select an analysis field"
                    End If
                
150             Case 5

152                 If Len(txtLayerName) > 0 And Len(txtDescription) > 0 And Len(cmbThemeGroup.Text) > 0 Then
154                     bVer = True
                    Else
156                     bVer = False
158                     MsgBox "Please enter in all detail"
                    End If
                
160             Case 6

162                 bVer = True

164             Case 7
                    
166                 If Len(txtIniFileName.fileName) > 0 Then
                        
168                     If bEditEntry = vbNo Then .AddNew

170                     With RSLocalThemes.fields
                
172                         .Item("Name").Value = txtThemeName
174                         .Item("ThemeGroup").Value = cmbThemeGroup.ItemData(cmbThemeGroup.ListIndex)
176                         .Item("Description").Value = txtDescription
178                         .Item("AnalysisField").Value = lstAnalysisFlds
180                         .Item("Maps").Value = txtMaps
182                         .Item("AnalysisLayer").Value = txtLayerName
184                         .Item("ThemeConfigName").Value = txtIniFileName
                
                        End With
                
186                     .CurrTab = 0

                    Else
188                     MsgBox "Please select an INI file!"
                    End If

190                 bVer = False
                
            End Select

192         If bVer Then
                
194             If Not .CurrTab = .NumTabs Then

196                 If .CurrTab = 1 Then
198                     If optGISSource(0).Value = True Or optGISSource(2).Value = True Then
200                         .CurrTab = .CurrTab + 1

202                         If optGISSource(2).Value = True Then
204                             txtLayerName.Text = m_sIncidentLayerName
                            End If
                        End If
                    End If
                    
206                 .CurrTab = .CurrTab + 1
                End If
                
            End If

        End With

        '<EhFooter>
        Exit Sub

cmdNext_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmThemeWiz.cmdNext_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub EditExistingEntry()

    Dim i As Integer
    optGISSource(0).Value = True
    FraGISAttribute.Enabled = False
    
    With RSLocalThemes.fields
                
        i = 0

        Do Until i = cmbThemeGroup.ListCount
        
            If cmbThemeGroup.ItemData(i) = .Item("ThemeGroup").Value Then
                cmbThemeGroup = cmbThemeGroup.List(i)
                cmbThemeGroup.ListIndex = i
            End If

            i = i + 1
        Loop
        
        txtThemeName = "" & .Item("Name").Value
        'cmbThemeGroup.Text = "" & .Item("ThemeGroup").Value
        txtDescription = "" & .Item("Description").Value
        lstAnalysisFlds = "" & .Item("AnalysisField").Value
        txtMaps = "" & .Item("Maps").Value
        txtLayerName = "" & .Item("AnalysisLayer").Value

        If m_sIncidentLayerName = txtLayerName.Text Then optGISSource(2).Value = True
        
        On Error GoTo nosuchfile
        txtIniFileName = "" & .Item("ThemeConfigName").Value
                
    End With

    Exit Sub

nosuchfile:
    MsgBox "File [" & RSLocalThemes.fields.Item("ThemeConfigName").Value & "] was not found"

End Sub

Private Sub cmdSave_Click()
        '<EhHeader>
        On Error GoTo cmdSave_Click_Err
        '</EhHeader>
        Dim bReturnValue As Boolean
    
100     RSLocalThemes.Filter = adFilterPendingRecords

102     If Not RSLocalThemes.EOF And Not RSLocalThemes.Bof Then
        
104         If MsgBox("Do you wish to save your changes?", vbYesNo, "Confirm Save") = vbYes Then
        
106             bReturnValue = m_frmOASISProgress.SaveHttpCommsRS(RSLocalThemes, WebSite & "Oasis.asp", True)
                
108             If bReturnValue Then

110                 IncrementProfileSettingVersion WebSite, "SettingValue10", RSLocalUserGroups.fields("Name").Value
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
               "in OASISRemoteAdmin.frmThemeWiz.cmdSave_Click " & _
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
    
100     c.DialogTitle = "Open Layer File"
102     c.Filter = "Layer Files (shp)|" & GisUtils.GisSupportedFiles(XgisFileTypeVector, False)
104     c.InitDir = CreateAppPath & "\data\user\maps"
106     c.ShowOpen

108     If c.fileName <> "" Then

110         SetLyrAttr c.fileName
112         Me.txtLayerName = c.fileName
        
114         txtLayerName = Right(txtLayerName, Len(txtLayerName) - InStrRev(txtLayerName, "\"))
116         txtLayerName = Left(txtLayerName, InStrRev(txtLayerName, ".") - 1)
        End If
         
        '<EhFooter>
        Exit Sub

cmdOpenGISFile_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmThemeWiz.cmdOpenGISFile_Click " & _
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
    
134     lstAnalysisFlds.Visible = True
136     lstAnalysisFlds.Clear
                        
138     With oLyr.fields
            
140         For i = 0 To .Count - 1
142             lstAnalysisFlds.AddItem .Item(i).Name
144             lstAnalysisFlds.Visible = True
            Next
            
        End With

        '<EhFooter>
        Exit Sub

SetLyrAttr_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmThemeWiz.SetLyrAttr " & _
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
         
110     lstAnalysisFlds.Visible = True
112     lstAnalysisFlds.Clear
                        
114     With RSIncidentsTable.fields
            
116         For i = 0 To .Count - 1
118             lstAnalysisFlds.AddItem .Item(i).Name
120             lstAnalysisFlds.Visible = True
            Next
            
        End With
        
122     Set RSIncidentsTable = Nothing
124     ooConn.Close

        '<EhFooter>
        Exit Sub

SetIncLyrAttr_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmThemeWiz.SetIncLyrAttr " & _
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
               "in OASISRemoteAdmin.frmThemeWiz.cmdTools_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDBGrid1_OnDblClick()
        '<EhHeader>
        On Error GoTo dxDBGrid1_OnDblClick_Err
        '</EhHeader>

100     If DeleteRecordFromRSAndSave(RSLocalThemes, "SettingValue10", RSLocalUserGroups.fields("Name").Value) Then Unload Me

        '<EhFooter>
        Exit Sub

dxDBGrid1_OnDblClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmThemeWiz.dxDBGrid1_OnDblClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub m_frmThemeGroups_RefreshThemeGroups()
        '<EhHeader>
        On Error GoTo m_frmThemeGroups_RefreshThemeGroups_Err
        '</EhHeader>

        Dim sString As String
    
100     sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & RSLocalUserGroups!Name & "ThemeGroups")
102     Set RSLocalThemeGroups = m_frmOASISProgress.OpenHttpCommsRS(sString, True)

104     If RSLocalThemeGroups.State = adStateClosed Then
106         MsgBox "ThemeGroups Table Does Not Exist!"
            Exit Sub
        End If
    
108     cmbThemeGroup.Clear
        
110     If Not RSLocalThemeGroups.EOF Or Not RSLocalThemeGroups.Bof Then
    
112         RSLocalThemeGroups.MoveFirst
     
114         Do Until RSLocalThemeGroups.EOF
     
116             Me.cmbThemeGroup.AddItem RSLocalThemeGroups!Name
118             Me.cmbThemeGroup.ItemData(Me.cmbThemeGroup.NewIndex) = RSLocalThemeGroups!id
120             RSLocalThemeGroups.MoveNext
     
            Loop
    
        End If
    
        '<EhFooter>
        Exit Sub

m_frmThemeGroups_RefreshThemeGroups_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmThemeWiz.m_frmThemeGroups_RefreshThemeGroups " & _
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
        
        C1TabMain.TabHeight = 1
        C1TabMain.CurrTab = 0
    
100     DoEvents
102     Me.Picture = g_PictureDialogSmall
104     Set RSLocalThemes = New ADODB.Recordset
106     Set RSLocalThemeGroups = New ADODB.Recordset

108     sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT SettingValue1 FROM " & RSLocalUserGroups!Name & "AppSettings WHERE SettingName = 'OASIS_Incident_Layer_Name'")
110     Set rsOName = m_frmOASISProgress.OpenHttpCommsRS(sString, True)
        
112     If Not rsOName.State = adStateClosed Then
114         If Not rsOName.Bof And Not rsOName.EOF Then
116             m_sIncidentLayerName = rsOName!SettingValue1
            End If
        End If
        
118     sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & RSLocalUserGroups!Name & "Themes")
120     Set RSLocalThemes = m_frmOASISProgress.OpenHttpCommsRS(sString, True)

122     If RSLocalThemes.State = adStateClosed Then
124         MsgBox "Themes Table Does Not Exist!"
            Exit Sub
        End If

126     Call m_frmThemeGroups_RefreshThemeGroups
    
128     txtIniFileName.Path = CreateAppPath & "\data\user\maps\Layers"

130     dxDBGrid1.Columns.DestroyColumns
        dxDBGrid1.KeyField = RSLocalThemes.fields(0).Name
132     Set dxDBGrid1.DataSource = RSLocalThemes
134     dxDBGrid1.Columns.RetrieveFields
136     dxDBGrid1.Columns(0).Visible = False
138     dxDBGrid1.Columns(1).Visible = True
140     dxDBGrid1.Columns(2).Visible = False
142     dxDBGrid1.Columns(3).Visible = False
144     dxDBGrid1.Columns(4).Visible = True
146     dxDBGrid1.Columns(5).Visible = False
148     dxDBGrid1.Columns(6).Visible = True
150     dxDBGrid1.Columns(7).Visible = True
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmThemeWiz.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

100     Set RSLocalThemes = Nothing
102     Set RSLocalThemeGroups = Nothing
104     Set RSLocalUserGroups = Nothing

106     Set m_frmThemeGroups = Nothing

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmThemeWiz.Form_Unload " & _
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
104         txtLayerName.Locked = True
106         txtLayerName.Enabled = True
        Else
108         txtLayerName.Locked = False
110         txtLayerName.Enabled = False
112         cmdOpenGISFile.Enabled = False
114         lstAnalysisFlds.Visible = True
        End If

        '<EhFooter>
        Exit Sub

optGISSource_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmThemeWiz.optGISSource_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
