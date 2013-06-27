VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Begin VB.UserControl OasisMAP 
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ScaleHeight     =   5820
   ScaleWidth      =   6765
   ToolboxBitmap   =   "OasisMAP.ctx":0000
   Begin C1SizerLibCtl.C1Elastic elTatukGIS 
      Height          =   5820
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6765
      _cx             =   11933
      _cy             =   10266
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
      _GridInfo       =   $"OasisMAP.ctx":0312
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin TatukGIS_DK.XGIS_ViewerWnd GIS 
         Height          =   5820
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6765
         BigExtentMargin =   -10
         RestrictedDrag  =   -1  'True
         CachedPaint     =   -1  'True
         IncrementalPaint=   -1  'True
         FullPaint       =   -1  'True
         CodePage        =   0
         OutCodePage     =   0
         CharSet         =   0
         UseRTree        =   0   'False
         PrinterTileSize =   512
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
         SelectionPattern=   "OasisMAP.ctx":0345
         SelectionTransparency=   100
         SelectionWidth  =   100
         SelectionOutlineOnly=   0   'False
         OldCachedPaint  =   0   'False
         PrinterModeDraft=   0   'False
         PrinterModeForceBitmap=   0   'False
         Mode            =   0
         BorderStyle     =   0
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
   End
End
Attribute VB_Name = "OasisMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Init()

End Sub
