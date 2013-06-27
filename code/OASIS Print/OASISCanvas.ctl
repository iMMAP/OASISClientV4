VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{9C989D2F-3596-477B-B719-6DC4E6893AF2}#1.0#0"; "TatukGIS_XDK10_WIN32.ocx"
Begin VB.UserControl OASISDrawObj 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   ToolboxBitmap   =   "OASISCanvas.ctx":0000
   Begin TatukGIS_XDK10.XGIS_ViewerWnd GIS 
      Height          =   6450
      Left            =   5580
      TabIndex        =   11
      Top             =   450
      Visible         =   0   'False
      Width           =   6630
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PrintFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      SelectionPattern=   "OASISCanvas.ctx":0312
      SelectionTransparency=   100
      SelectionWidth  =   100
      SelectionOutlineOnly=   0   'False
      OldCachedPaint  =   0   'False
      PrinterModeDraft=   0   'False
      PrinterModeForceBitmap=   0   'False
      GDIType         =   1
      ScaleAsFloat    =   1
      Mode            =   2
      BorderStyle     =   1
      CursorForDrag   =   0
      CursorForSelect =   0
      CursorForZoom   =   0
      CursorForEdit   =   0
      MinZoomSize     =   -5
      ScrollBars      =   0
      AutoCenter      =   0   'False
      Align           =   0
      Ctl3D           =   0   'False
      Object.Visible         =   -1  'True
      Cursor          =   18
      DoubleBuffered  =   0   'False
      ModeMouseButton =   0
      CursorForUserDefined=   0
      View3D          =   0   'False
   End
   Begin TatukGIS_XDK10.XGIS_ControlNorthArrow NorthArrow 
      Height          =   465
      Left            =   3960
      TabIndex        =   16
      Top             =   855
      Visible         =   0   'False
      Width           =   555
      Symbol          =   0
      Transparent     =   0   'False
      Path            =   ""
      Align           =   0
      BevelInner      =   0
      BevelOuter      =   2
      BorderStyle     =   0
      Color           =   -16777201
      Ctl3D           =   -1  'True
      Enabled         =   -1  'True
      FullRepaint     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColor       =   -2147483630
      Object.Visible         =   -1  'True
      DoubleBuffered  =   -1  'True
      Color2          =   0
      Color1          =   0
   End
   Begin TatukGIS_XDK10.XGIS_ControlScale ScaleBar 
      Height          =   285
      Left            =   1305
      TabIndex        =   17
      Top             =   3150
      Visible         =   0   'False
      Width           =   2040
      Dividers        =   5
      Align           =   0
      BevelInner      =   0
      BevelOuter      =   2
      BorderStyle     =   0
      Color           =   -16777201
      Ctl3D           =   -1  'True
      Enabled         =   -1  'True
      FullRepaint     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColor       =   -2147483630
      Object.Visible         =   -1  'True
      DoubleBuffered  =   0   'False
      UnitsEPSG       =   904201
   End
   Begin TatukGIS_XDK10.XGIS_ControlLegend Legend 
      Height          =   1455
      Left            =   0
      TabIndex        =   12
      Top             =   1260
      Visible         =   0   'False
      Width           =   915
      BorderStyle     =   1
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
         Size            =   8.25
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
      Ctl3D           =   -1  'True
      Color           =   -2147483633
      Enabled         =   -1  'True
      Object.Visible         =   -1  'True
      DoubleBuffered  =   -1  'True
      AllowMove       =   -1  'True
      AllowActive     =   0   'False
      AllowExpand     =   -1  'True
      AllowParams     =   -1  'True
      Mode            =   0
   End
   Begin MSComctlLib.Toolbar tbrMapTools 
      Height          =   330
      Left            =   2475
      TabIndex        =   15
      Top             =   2340
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "zoommap"
            Object.ToolTipText     =   "Zoom in the map"
            Object.Tag             =   "zoommap"
            ImageIndex      =   35
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "panmap"
            Object.ToolTipText     =   "Move the map center"
            Object.Tag             =   "panmap"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "closemap"
            Object.ToolTipText     =   "Close Interactive Map"
            Object.Tag             =   "closemap"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "mapsettings"
            Object.ToolTipText     =   "Map settings"
            Object.Tag             =   "mapsettings"
            ImageIndex      =   37
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Height          =   240
      Left            =   4995
      ScaleHeight     =   180
      ScaleWidth      =   360
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3060
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   4995
      ScaleHeight     =   225
      ScaleWidth      =   315
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2565
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   825
      Left            =   225
      ScaleHeight     =   765
      ScaleWidth      =   945
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   50
      Left            =   1890
      Max             =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3420
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2085
      LargeChange     =   50
      Left            =   4545
      Max             =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1305
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton Corner 
      Height          =   240
      Left            =   4545
      TabIndex        =   5
      Top             =   3420
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicData 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   4950
      ScaleHeight     =   825
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox DrawControl 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawMode        =   6  'Mask Pen Not
      Height          =   2355
      Left            =   180
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   1
      Top             =   720
      Width           =   3615
      Begin VB.TextBox myText 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   480
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "OASISCanvas.ctx":1FC4
         Top             =   690
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label griff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   0
         Left            =   60
         MousePointer    =   2  'Cross
         TabIndex        =   9
         Top             =   90
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label griff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   1
         Left            =   240
         MousePointer    =   2  'Cross
         TabIndex        =   8
         Top             =   90
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label griff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   2
         Left            =   420
         MousePointer    =   2  'Cross
         TabIndex        =   7
         Top             =   90
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label griff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   3
         Left            =   600
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   90
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   405
      Top             =   3285
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   37
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":1FCE
            Key             =   "New"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":20E0
            Key             =   "Open"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":21F2
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":2304
            Key             =   "Export"
            Object.Tag             =   "Export"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":2656
            Key             =   "Cut"
            Object.Tag             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":2768
            Key             =   "Copy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":287A
            Key             =   "Paste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":298C
            Key             =   "Undo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":2A9E
            Key             =   "Redo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":2BB0
            Key             =   "Delete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":2CC2
            Key             =   "TextLeft"
            Object.Tag             =   "TextLeft"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":2DD4
            Key             =   "TextCenter"
            Object.Tag             =   "TextCenter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":2EE6
            Key             =   "TextRight"
            Object.Tag             =   "TextRight"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":2FF8
            Key             =   "Bold"
            Object.Tag             =   "Bold"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":310A
            Key             =   "Italic"
            Object.Tag             =   "Italic"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":321C
            Key             =   "Underline"
            Object.Tag             =   "Underline"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":332E
            Key             =   "Strikethru"
            Object.Tag             =   "Strikethru"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":3440
            Key             =   "SelectAll"
            Object.Tag             =   "SelectAll"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":3792
            Key             =   "UnselectAll"
            Object.Tag             =   "UnselectAll"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":3AE4
            Key             =   "AlignLeft"
            Object.Tag             =   "AlignLeft"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":3E36
            Key             =   "AlignCenterVertical"
            Object.Tag             =   "AlignCenterVertical"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":4188
            Key             =   "AlignRight"
            Object.Tag             =   "AlignRight"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":44DA
            Key             =   "AlignTop"
            Object.Tag             =   "AlignTop"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":482C
            Key             =   "AlignCenterHorizontal"
            Object.Tag             =   "AlignCenterHorizontal"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":4B7E
            Key             =   "AlignBottom"
            Object.Tag             =   "AlignBottom"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":4ED0
            Key             =   "AlignCenterVerticalHorizontal"
            Object.Tag             =   "AlignCenterVerticalHorizontal"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":5222
            Key             =   "BringToFront"
            Object.Tag             =   "BringToFront"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":5334
            Key             =   "SendToBack"
            Object.Tag             =   "SendToBack"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":5446
            Key             =   "BringForward"
            Object.Tag             =   "BringForward"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":5558
            Key             =   "SendBackward"
            Object.Tag             =   "SendBackward"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":566A
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":577C
            Key             =   "Ungroup"
            Object.Tag             =   "Ungroup"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":588E
            Key             =   "Zoom100"
            Object.Tag             =   "Zoom100"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":5BE0
            Key             =   "Zoom-"
            Object.Tag             =   "Zoom-"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":5F32
            Key             =   "Zoom+"
            Object.Tag             =   "Zoom+"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":6284
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OASISCanvas.ctx":63B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image cRoundRect 
      Height          =   480
      Left            =   4650
      Picture         =   "OASISCanvas.ctx":6710
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cStar 
      Height          =   480
      Left            =   3900
      Picture         =   "OASISCanvas.ctx":6A1A
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cPolygon 
      Height          =   480
      Left            =   3330
      Picture         =   "OASISCanvas.ctx":6D24
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cArc 
      Height          =   480
      Left            =   2730
      Picture         =   "OASISCanvas.ctx":702E
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cPicture 
      Height          =   480
      Left            =   2160
      Picture         =   "OASISCanvas.ctx":7338
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cText 
      Height          =   480
      Left            =   1650
      Picture         =   "OASISCanvas.ctx":7642
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cEllipse 
      Height          =   480
      Left            =   1140
      Picture         =   "OASISCanvas.ctx":794C
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cRect 
      Height          =   480
      Left            =   600
      Picture         =   "OASISCanvas.ctx":7C56
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cLine 
      Height          =   480
      Left            =   60
      Picture         =   "OASISCanvas.ctx":7F60
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "OASISDrawObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lLastMouseX As Long
Private lLastMouseY As Long

Private Type myObject
    mObjectType As myObType
    mTop As Single
    mLeft As Single
    mHeight As Single
    mWidth As Single
    mAngle As Single
    mFillColor As Long
    mFillStyle As myFill
    mBorderColor As Long
    mBorderWidth As Integer
    mAspect As Single
    mDeleted As Boolean
    mPicture As StdPicture
    mFontName As String
    mFontSize As Single
    mFontBold As Boolean
    mFontItalic As Boolean
    mFontUnderline As Boolean
    mFontStrikethru As Boolean
    mText As String
    mTextAlign As AlignmentConstants
    mPointQty As Integer
    mPosX0 As Single
    mPosY0 As Single
    mPosX1 As Single
    mPosY1 As Single
    mPosX2 As Single
    mPosY2 As Single
    mPosX3 As Single
    mPosY3 As Single
    mGroupMember As Integer
    mGISObj As myGISobj
    'mGISPath As String
End Type

Public Enum myGISobj
    oNone = 0
    oMap = 1
    oLegend = 2
    oNortArrow = 3
    oScaleBar = 4
    oDataGrid = 5
End Enum

Private Type myCoorType
    posX1 As Single
    posY1 As Single
    posX2 As Single
    posY2 As Single
    posX3 As Single
    posY3 As Single
    posX4 As Single
    posY4 As Single
    centerX As Single
    centerY As Single
End Type

Public Enum myObType
    mline = 0
    mArc = 1
    mRectangle = 2
    mEllipse = 3
    mText = 4
    mImage = 5
    mPolygon = 6
    mStar = 7
    mRoundRectangle = 8
End Enum

Public Enum myOrder
    BringToFront = 0
    SendToBack = 1
    BringFoward = 2
    SendBackward = 3
End Enum

Public Enum myBool3
    Unchanged = -1
    mFalse = 0
    mTrue = 1
End Enum

Public Enum myFill
    mSolid = 0
    mTransparent = 1
    mHorizontalLine = 2
    mVerticalLine = 3
    mUpwardDiagonal = 4
    mDownwardDiagonal = 5
    mCross = 6
    mDiagonalCross = 7
End Enum

Public Enum myAlignType
    mLeft = 0
    mCenter_V = 1
    mRight = 2
    mTop = 3
    mCenter_H = 4
    mBottom = 5
    mCenter_V_H = 6
    mCenterPage = 7
End Enum

Public Enum myChange
    toLeft = 0
    toTop = 1
    toRight = 2
    toBottom = 3
    toWidthP = 4
    toHeightP = 5
    toWidthN = 6
    toHeightN = 7
End Enum

Dim ObjList()       As myObject
Dim ObjQty          As Long
Dim ObjIndex        As Long
Dim NewObj          As Boolean
Dim isDown          As Boolean
Dim isMove          As Boolean
Dim onObject        As Boolean
Dim isResize        As Boolean
Dim toSize          As Integer
Dim NextLine        As Boolean
Dim NewText         As Boolean
Dim mArrowStep      As Integer
Dim myFont          As String

Dim ListSel()       As Long
Dim QtySel          As Long

Dim Xmove           As Single
Dim Ymove           As Single

Dim LeftSel         As Single
Dim TopSel          As Single

Dim DownX           As Single
Dim DownY           As Single
Dim MouseSel        As Boolean

Dim mClipBoard()    As myObject
Dim ClipQty         As Long

Dim UndoBuffer()    As String
Dim mUndoSize       As Integer
Dim UndoStack       As Integer
Dim UndoPointer     As Integer
Dim isUndo          As Boolean

Dim mDefaultText    As String
Dim mCanvasWidth    As Long
Dim mCanvasHeight   As Long
Dim mCanvasCenterX   As Long
Dim mCanvasCenterY   As Long
Dim mShowCanvasSize As Boolean
Dim mZF             As Single
Dim toZoom          As Boolean
Dim GroupQty        As Integer

Dim Drag            As Boolean

Private Const Pi = 3.14159265358979

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyAscii As Integer, Shift As Integer)
Public Event KeyPress(KeyCode As Integer)
Public Event KeyUp(KeyAscii As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event NewDrawingEnd()

Public Event UndoRedo(LastUndo As Boolean, LastRedo As Boolean)
Public Event Prompt2Save()

Public Event ObjSelected(ObjType As myObType, oGISObj As myGISobj, _
                         Index As Long, _
                         ObjLeft As Single, _
                         ObjTop As Single, _
                         ObjWidth As Single, _
                         ObjHeight As Single, _
                         ObjAngle As Single, _
                         ObjFillColor As Long, _
                         ObjFillStyle As myFill, _
                         ObjBorderColor As Long, _
                         ObjBorderWidth As Integer, _
                         ObjAspect As Single, _
                         ObjFontName As String, _
                         ObjFontSize As Single, _
                         ObjFontBold As Boolean, _
                         ObjFontItalic As Boolean, _
                         ObjFontUnderline As Boolean, _
                         ObjFontStrikethru As Boolean, _
                         ObjText As String, _
                         ObjTextAlign As AlignmentConstants, _
                         ObjPointQty As Integer)

Public Event ObjectResize(ObjType As myObType, _
                          Index As Long, _
                          ObjLeft As Single, _
                          ObjTop As Single, _
                          ObjWidth As Single, _
                          ObjHeight As Single, _
                          ObjAspect As Single)

Private Declare Function ExtFloodFill _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal crColor As Long, _
                             ByVal wFillType As Long) As Long
Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

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

Private Declare Function CreateFont _
                Lib "gdi32.dll" _
                Alias "CreateFontA" (ByVal nHeight As Long, _
                                     ByVal nWidth As Long, _
                                     ByVal nEscapement As Long, _
                                     ByVal nOrientation As Long, _
                                     ByVal fnWeight As Long, _
                                     ByVal fdwItalic As Long, _
                                     ByVal fdwUnderline As Long, _
                                     ByVal fdwStrikeOut As Long, _
                                     ByVal fdwCharSet As Long, _
                                     ByVal fdwOutputPrecision As Long, _
                                     ByVal fdwClipPrecision As Long, _
                                     ByVal fdwQuality As Long, _
                                     ByVal fdwPitchAndFamily As Long, _
                                     ByVal lpszFace As String) As Long
Private Declare Function GetDeviceCaps _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal nIndex As Long) As Long
Private Declare Function MulDiv _
                Lib "kernel32" (ByVal nNumber As Long, _
                                ByVal nNumerator As Long, _
                                ByVal nDenominator As Long) As Long

Private Const LOGPIXELSY = 90                    'For GetDeviceCaps - returns the height of a logical pixel
Private Const ANSI_CHARSET = 0                   'Use the default Character set
Private Const CLIP_LH_ANGLES = 16                ' Needed for tilted fonts.
Private Const OUT_TT_PRECIS = 9                  'Tell it to use True Types when Possible
Private Const PROOF_QUALITY = 9                  'Make it as clean as possible.
Private Const DEFAULT_PITCH = 0                  'We want the font to take whatever pitch it defaults to
Private Const FF_DONTCARE = 0                    'Use whatever fontface it is.

Private Enum FontWeight
    FW_DONTCARE = 0
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_ULTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_REGULAR = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_DEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_ULTRABOLD = 800
    FW_HEAVY = 900
    FW_BLACK = 900
End Enum

Private Declare Function PolyBezier _
                Lib "gdi32" (ByVal hdc As Long, _
                             lpPoint As POINTAPI, _
                             ByVal nCount As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function Polygon _
                Lib "gdi32" (ByVal hdc As Long, _
                             lpPoint As POINTAPI, _
                             ByVal nCount As Long) As Long
Private Declare Function RoundRect _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal X1 As Long, _
                             ByVal Y1 As Long, _
                             ByVal X2 As Long, _
                             ByVal Y2 As Long, _
                             ByVal X3 As Long, _
                             ByVal Y3 As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function PlgBlt _
                Lib "gdi32" (ByVal hdcDest As Long, _
                             lpPoint As POINTAPI, _
                             ByVal hdcSrc As Long, _
                             ByVal nXSrc As Long, _
                             ByVal nYSrc As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hbmMask As Long, _
                             ByVal xMask As Long, _
                             ByVal yMask As Long) As Long

Dim GisUtils As New XGIS_Utils

Private Sub SuperDebug(sText As String)
    DebugPrint "Calling OASISDrawObj." & sText, True
End Sub

Private Function EllipsePts(cLeft As Single, _
                            cTop As Single, _
                            cWidth As Single, _
                            cHeight As Single, _
                            cAngle As Single) As POINTAPI()
    SuperDebug "sub/fun: EllipsePts"
    
    Dim offsetX   As Single
    Dim offsetY   As Single
    Dim R         As Single
    Dim Alfa      As Single
    Dim px(12)    As Single
    Dim py(12)    As Single
    Dim POINT(12) As POINTAPI
    Dim n         As Integer
    Dim centerX   As Single
    Dim centerY   As Single
    Dim eFactor   As Double

    eFactor = 2 / 3 * (Sqr(2) - 1)

    offsetX = cWidth * eFactor
    offsetY = cHeight * eFactor
    centerX = cWidth / 2
    centerY = cHeight / 2
    
    px(0) = cWidth
    px(1) = px(0)
    px(11) = px(0)
    px(12) = px(0)
    
    px(5) = 0
    px(6) = px(5)
    px(7) = px(5)
    
    px(2) = centerX + offsetX
    px(10) = px(2)
    
    px(4) = centerX - offsetX
    px(8) = px(4)

    px(3) = centerX
    px(9) = px(3)
    
    py(2) = 0
    py(3) = py(2)
    py(4) = py(2)
    
    py(8) = cHeight
    py(9) = py(8)
    py(10) = py(8)
    
    py(7) = centerY + offsetY
    py(11) = py(7)
    
    py(1) = centerY - offsetY
    py(5) = py(1)
    
    py(0) = centerY
    py(12) = py(0)
    py(6) = py(0)

    For n = 0 To 12
        R = Sqr(px(n) ^ 2 + py(n) ^ 2)
        Alfa = Atn2(py(n), px(n)) - (cAngle * Pi / 180)
        POINT(n).X = cLeft + R * Cos(Alfa)
        POINT(n).Y = cTop + R * Sin(Alfa)
    Next n

    EllipsePts = POINT
End Function

Private Function Atn2(ByVal Y As Single, ByVal X As Single) As Single

    SuperDebug "sub/fun: Atn2"
    If X = 0 Then
        Atn2 = IIf(Y = 0, Pi / 4, Sgn(Y) * Pi / 2)
    Else
        Atn2 = Atn(Y / X) + (1 - Sgn(X)) * Pi / 2
    End If

End Function

Public Sub AddObject(tObjectType As myObType, _
                     Optional tTop As Single = -1, _
                     Optional tLeft As Single = -1, _
                     Optional tHeight As Single = -1, _
                     Optional tWidth As Single = -1, _
                     Optional tAngle As Single, _
                     Optional tFillColor As Long = -1, _
                     Optional tFillStyle As myFill = mSolid, _
                     Optional tBorderColor As Long = -1, _
                     Optional tBorderWidth As Integer = 0, _
                     Optional tPicture As StdPicture, _
                     Optional tFontName As String = "", _
                     Optional tFontSize As Single = 8, _
                     Optional tFontBold As Boolean = False, _
                     Optional tFontItalic As Boolean = False, _
                     Optional tFontUnderline As Boolean = False, _
                     Optional tFontStrikethru As Boolean = False, _
                     Optional tText As String = "", _
                     Optional tTextAlign As AlignmentConstants = vbLeftJustify, _
                     Optional tPointQty As Integer = 3, _
                     Optional tPosX0 As Single = -1, _
                     Optional tPosY0 As Single = -1, _
                     Optional tPosX1 As Single = -1, _
                     Optional tPosY1 As Single = -1, _
                     Optional tPosX2 As Single = -1, Optional tPosY2 As Single = -1, Optional tPosX3 As Single = -1, Optional tPosY3 As Single = -1, Optional tGroupMember As Integer = 0, Optional tAspect As Single, Optional tGISObj As myGISobj = oNone)

    SuperDebug "sub/fun: AddObject"
    NextLine = False

    NewObj = False

    Add2Selection -1

    If tObjectType = mText Then
        If tText = "" Then tText = mDefaultText
        If tFontName = "" Then tFontName = myFont
    Else
        tText = ""
        tFontName = ""
        tFontSize = 0
        tFontBold = False
        tFontItalic = False
        tFontUnderline = False
        tFontStrikethru = False
    End If
    
    On Error GoTo eric
    
    ObjQty = ObjQty + 1
    ReDim Preserve ObjList(ObjQty)
    ObjIndex = ObjQty - 1

    Add2Selection ObjIndex

    With ObjList(ObjIndex)
        .mObjectType = tObjectType
        .mTop = tTop
        .mLeft = tLeft
        .mHeight = tHeight
        .mWidth = tWidth
        .mAngle = tAngle

        If .mObjectType = mArc Then .mAngle = 0
        .mFillColor = tFillColor
        .mFillStyle = tFillStyle
        .mBorderColor = tBorderColor
        .mBorderWidth = tBorderWidth
        .mFontName = tFontName
        .mFontSize = tFontSize
        .mFontBold = tFontBold
        .mFontItalic = tFontItalic
        .mFontUnderline = tFontUnderline
        .mFontStrikethru = tFontStrikethru
        .mText = tText
        .mTextAlign = tTextAlign
        .mPointQty = tPointQty
        .mPosX0 = tPosX0
        .mPosY0 = tPosY0
        .mPosX1 = tPosX1
        .mPosY1 = tPosY1
        .mPosX2 = tPosX2
        .mPosY2 = tPosY2
        .mPosX3 = tPosX3
        .mPosY3 = tPosY3
        .mGroupMember = tGroupMember
        .mAspect = tAspect
        Set .mPicture = tPicture
        .mGISObj = tGISObj
    End With

    If tTop = -1 And tLeft = -1 Then ' if no position mouse position will be used

        Select Case tObjectType

            Case mline
                DrawControl.MousePointer = 99
                Set DrawControl.MouseIcon = cLine.Picture

            Case mArc
                DrawControl.MousePointer = 99
                Set DrawControl.MouseIcon = cArc.Picture

            Case mRectangle
                DrawControl.MousePointer = 99
                Set DrawControl.MouseIcon = cRect.Picture

            Case mRoundRectangle
                DrawControl.MousePointer = 99
                Set DrawControl.MouseIcon = cRoundRect.Picture

            Case mEllipse
                DrawControl.MousePointer = 99
                Set DrawControl.MouseIcon = cEllipse.Picture

            Case mText
                DrawControl.MousePointer = 99
                Set DrawControl.MouseIcon = cText.Picture
                DrawControl.Font = ObjList(ObjIndex).mFontName
                DrawControl.FontSize = ObjList(ObjIndex).mFontSize
                DrawControl.FontBold = ObjList(ObjIndex).mFontBold
                DrawControl.FontItalic = ObjList(ObjIndex).mFontItalic
                DrawControl.FontUnderline = ObjList(ObjIndex).mFontUnderline
                DrawControl.FontStrikethru = ObjList(ObjIndex).mFontStrikethru
                ObjList(ObjIndex).mWidth = DrawControl.TextWidth(tText) + DrawControl.TextWidth("XX")
                ObjList(ObjIndex).mHeight = DrawControl.TextHeight(tText)
                NewText = True

            Case mImage
                DrawControl.MousePointer = 99
                Set DrawControl.MouseIcon = cPicture.Picture

            Case mPolygon
                DrawControl.MousePointer = 99
                Set DrawControl.MouseIcon = cPolygon.Picture

            Case mStar
                DrawControl.MousePointer = 99
                Set DrawControl.MouseIcon = cStar.Picture
        End Select

        NewObj = True
    Else
        DrawControl.MousePointer = 0

        Select Case tObjectType

            Case mText
                DrawControl.Font = ObjList(ObjIndex).mFontName
                DrawControl.FontSize = ObjList(ObjIndex).mFontSize
                DrawControl.FontBold = ObjList(ObjIndex).mFontBold
                DrawControl.FontItalic = ObjList(ObjIndex).mFontItalic
                DrawControl.FontUnderline = ObjList(ObjIndex).mFontUnderline
                DrawControl.FontStrikethru = ObjList(ObjIndex).mFontStrikethru

                If tWidth = -1 Or tHeight = -1 Then
                    ObjList(ObjIndex).mWidth = DrawControl.TextWidth(tText) + DrawControl.TextWidth("XX")
                    ObjList(ObjIndex).mHeight = DrawControl.TextHeight(tText)
                End If

            Case mImage

                If ObjList(ObjIndex).mWidth = -1 Then
                    ObjList(ObjIndex).mWidth = DrawControl.ScaleX(ObjList(ObjIndex).mPicture.Width)
                End If

                If ObjList(ObjIndex).mHeight = -1 Then
                    ObjList(ObjIndex).mHeight = DrawControl.ScaleY(ObjList(ObjIndex).mPicture.Height)
                End If

        End Select

        RaiseEvent NewDrawingEnd
        ReDraw
    End If

eric:

End Sub

Public Property Get CurrentObject() As Long
    CurrentObject = ObjIndex
End Property

Public Property Get ObjectInClipboard() As Boolean
    ObjectInClipboard = CBool(ClipQty)
End Property

Public Property Get Image() As StdPicture
    Set Image = DrawControl.Image
End Property

Public Property Get ObjectType() As myObType
    On Error Resume Next
    ObjectType = ObjList(ObjIndex).mObjectType
End Property

Public Property Get ObjectQty() As Long
    ObjectQty = ObjQty
End Property

Public Property Get SelectionQty() As Long
    SelectionQty = QtySel
End Property

Private Sub Command1_Click()
    SuperDebug "sub/fun: Command1_Click"
    DrawGIS PicData
End Sub

Private Sub Corner_Click()
    SuperDebug "sub/fun: Corner_Click"
    HScroll1.value = (HScroll1.Max - HScroll1.Min) \ 2
    VScroll1.value = (VScroll1.Max - VScroll1.Min) \ 2
End Sub

Private Sub DrawControl_DragDrop(Source As Control, X As Single, Y As Single)
    SuperDebug "sub/fun: DrawControl_DragDrop"
    DragBezier Source.Index, X, Y
    Add2UndoBuffer
End Sub

Private Sub DrawControl_DragOver(Source As Control, _
                                 X As Single, _
                                 Y As Single, _
                                 State As Integer)
    SuperDebug "sub/fun: DrawControl_DragOver"
    DragBezier Source.Index, X, Y
End Sub

Private Sub GIS_OnDblClick(translated As Boolean)
    SuperDebug "sub/fun: GIS_OnDblClick"
    GIS.Visible = False
    tbrMapTools.Visible = False
    ReDraw
End Sub

Private Sub griff_MouseMove(Index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

    SuperDebug "sub/fun: griff_MouseMove"
    If Button = 1 Then
        griff(Index).Drag
        Drag = True
    End If

End Sub

Private Sub griff_MouseUp(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)
    SuperDebug "sub/fun: griff_MouseUp"
    Drag = False
End Sub

Private Sub HScroll1_Change()
    On Error Resume Next
    SuperDebug "sub/fun: HScroll1_Change"
    
    DrawControl.left = HScroll1.value
    UserControl.Cls
    UserControl.DrawWidth = 1
    UserControl.Line (DrawControl.left + 4, DrawControl.top + 4)-Step(DrawControl.Width + 2, DrawControl.Height + 2), &H80000015, BF
    UserControl.Line (DrawControl.left - 1, DrawControl.top - 1)-Step(DrawControl.Width + 1, DrawControl.Height + 1), , B

    If mShowCanvasSize = True Then
        UserControl.CurrentX = DrawControl.left + DrawControl.Width - UserControl.TextWidth(mCanvasWidth & " X " & mCanvasHeight) + 7
        UserControl.CurrentY = DrawControl.top + DrawControl.Height + 7
        UserControl.Print mCanvasWidth & " X " & mCanvasHeight
    End If

    DrawControl.SetFocus
End Sub

Private Sub DrawControl_Click()
    SuperDebug "sub/fun: DrawControl_Click"
    RaiseEvent Click
End Sub

Private Sub DrawControl_DblClick()
    Dim n As Integer
    ' for edit selected text and arc

    SuperDebug "sub/fun: DrawControl_DblClick"
    If ObjIndex = -1 Then Exit Sub

    If ObjList(ObjIndex).mObjectType = mText Then
        NewText = True
        DrawControl.Font = ObjList(ObjIndex).mFontName
        DrawControl.FontSize = ObjList(ObjIndex).mFontSize
        DrawControl.FontBold = ObjList(ObjIndex).mFontBold
        DrawControl.FontItalic = ObjList(ObjIndex).mFontItalic
        DrawControl.FontUnderline = ObjList(ObjIndex).mFontUnderline
        DrawControl.FontStrikethru = ObjList(ObjIndex).mFontStrikethru
        myText.left = ObjList(ObjIndex).mLeft * mZF
        myText.top = ObjList(ObjIndex).mTop * mZF
        myText.Font = ObjList(ObjIndex).mFontName
        myText.FontSize = ObjList(ObjIndex).mFontSize * mZF
        myText.FontBold = ObjList(ObjIndex).mFontBold
        myText.FontItalic = ObjList(ObjIndex).mFontItalic
        myText.FontUnderline = ObjList(ObjIndex).mFontUnderline
        myText.FontStrikethru = ObjList(ObjIndex).mFontStrikethru
        myText.Text = ObjList(ObjIndex).mText
        myText.Width = ObjList(ObjIndex).mWidth * mZF
        myText.Height = ObjList(ObjIndex).mHeight * mZF
        myText.Visible = True
        ObjList(ObjIndex).mText = ""
        ReDraw
        myText.SelStart = 0
        myText.SelLength = Len(myText.Text)
        myText.SetFocus

    ElseIf ObjList(ObjIndex).mObjectType = mArc Then
        ReDraw False
        griff(0).left = ObjList(ObjIndex).mPosX0 * mZF
        griff(0).top = ObjList(ObjIndex).mPosY0 * mZF
        griff(1).left = ObjList(ObjIndex).mPosX1 * mZF
        griff(1).top = ObjList(ObjIndex).mPosY1 * mZF
        griff(2).left = ObjList(ObjIndex).mPosX2 * mZF
        griff(2).top = ObjList(ObjIndex).mPosY2 * mZF
        griff(3).left = ObjList(ObjIndex).mPosX3 * mZF
        griff(3).top = ObjList(ObjIndex).mPosY3 * mZF
        DrawControl.DrawStyle = vbDot
        DrawControl.DrawMode = vbInvert
        DrawControl.Line (griff(0).left + 4, griff(0).top)-(griff(1).left + 4, griff(1).top)
        DrawControl.Line (griff(2).left + 4, griff(2).top)-(griff(3).left + 4, griff(3).top)
        DrawControl.DrawStyle = vbSolid
        DrawControl.DrawMode = 13

        For n = 0 To 3
            griff(n).Visible = True
        Next n
    
    ElseIf ObjList(ObjIndex).mObjectType = mImage And ObjList(ObjIndex).mGISObj = oMap Then
        GIS.Move ObjList(ObjIndex).mLeft + DrawControl.left, ObjList(ObjIndex).mTop + DrawControl.top, ObjList(ObjIndex).mWidth, ObjList(ObjIndex).mHeight
        If Not GIS.Visible Then GIS.Visible = True
        tbrMapTools.Move GIS.left, GIS.top
        tbrMapTools.Visible = True
        tbrMapTools.ZOrder 0
        SetMapToolBar "zoommap"
        tbrMapTools.Refresh
    ElseIf ObjList(ObjIndex).mObjectType = mImage And ObjList(ObjIndex).mGISObj = oLegend Then
        Legend.Move ObjList(ObjIndex).mLeft + DrawControl.left, ObjList(ObjIndex).mTop + DrawControl.top, ObjList(ObjIndex).mWidth, ObjList(ObjIndex).mHeight
        If Not Legend.Visible Then Legend.Visible = True
    End If

    RaiseEvent DblClick
End Sub

Private Sub Legend_OnDblClick(translated As Boolean)
    SuperDebug "sub/fun: Legend_OnDblClick"
    Legend.Visible = False
    
End Sub

Private Sub tbrMapTools_ButtonClick(ByVal Button As MSComctlLib.Button)
    SuperDebug "sub/fun: tbrMapTools_ButtonClick"
    SetMapToolBar Button.Tag
End Sub

Private Sub SetMapToolBar(sMode As String)
    SuperDebug "sub/fun: SetMapToolBar"
    Dim i As Integer
    
    For i = 1 To tbrMapTools.Buttons.Count
        If tbrMapTools.Buttons.Item(i).Tag = sMode Then
            tbrMapTools.Buttons.Item(i).value = tbrPressed
        Else
            tbrMapTools.Buttons.Item(i).value = tbrUnpressed
        End If
    Next
    
    tbrMapTools.Refresh
    
    Select Case sMode
    
        Case "zoommap"
            GIS.Mode = XgisZoom
        Case "panmap"
            GIS.Mode = XgisDrag
        Case "closemap"
            GIS.Visible = False
            tbrMapTools.Visible = False
            ReDraw
        Case "mapsettings"
        
        
    End Select
End Sub

Private Sub UserControl_Click()

    SuperDebug "sub/fun: UserControl_Click"
    If NewObj = False And NewText = False And NextLine = False Then
        RaiseEvent ObjSelected(-1, oNone, -1, -1, -1, -1, -1, 0, -1, 0, -1, -1, -1, -1, -1, False, False, False, False, -1, -1, -1)
        ObjIndex = -1
        QtySel = 0
        ReDraw
    End If

End Sub

Private Sub UserControl_GotFocus()
    SuperDebug "sub/fun: UserControl_GotFocus"
    DrawControl.SetFocus
End Sub

Private Sub UserControl_Initialize()
    ObjIndex = -1
    myFont = "Arial"
End Sub

Private Sub UserControl_InitProperties()
    mDefaultText = "New Text"
End Sub

Private Sub DrawControl_KeyDown(KeyCode As Integer, Shift As Integer)
    'used for arrow keys
    SuperDebug "sub/fun: DrawControl_KeyDown"
    Dim n As Long

    Select Case Shift

        Case 0

            Select Case KeyCode

                Case vbKeyAdd
                    mZF = mZF + 0.1

                    If mZF > 10 Then mZF = 10
                    toZoom = True
                    UserControl_Resize
                    ReDraw

                Case vbKeySubtract
                    mZF = mZF - 0.1

                    If mZF < 0.1 Then mZF = 0.1
                    toZoom = True
                    UserControl_Resize
                    ReDraw
            End Select

            If QtySel > 0 Then

                For n = 0 To QtySel - 1

                    Select Case KeyCode

                        Case vbKeyLeft
                            ObjList(ListSel(n)).mLeft = ObjList(ListSel(n)).mLeft - 1 * mArrowStep * mZF

                            If ObjList(ListSel(n)).mObjectType = mArc Then EditArc ListSel(n), toLeft, mArrowStep * mZF

                        Case vbKeyUp
                            ObjList(ListSel(n)).mTop = ObjList(ListSel(n)).mTop - 1 * mArrowStep * mZF

                            If ObjList(ListSel(n)).mObjectType = mArc Then EditArc ListSel(n), toTop, mArrowStep * mZF

                        Case vbKeyRight
                            ObjList(ListSel(n)).mLeft = ObjList(ListSel(n)).mLeft + 1 * mArrowStep * mZF

                            If ObjList(ListSel(n)).mObjectType = mArc Then EditArc ListSel(n), toRight, mArrowStep * mZF

                        Case vbKeyDown
                            ObjList(ListSel(n)).mTop = ObjList(ListSel(n)).mTop + 1 * mArrowStep * mZF

                            If ObjList(ListSel(n)).mObjectType = mArc Then EditArc ListSel(n), toBottom, mArrowStep * mZF
                    End Select

                Next n

                ReDraw
            End If

        Case 2

            Select Case KeyCode

                Case vbKeyLeft
                    ObjIndex = ObjIndex - 1

                Case vbKeyUp
                    ObjIndex = ObjIndex + 1

                Case vbKeyRight
                    ObjIndex = ObjIndex + 1

                Case vbKeyDown
                    ObjIndex = ObjIndex - 1
            End Select

            If ObjIndex <= -1 Then ObjIndex = ObjQty - 1
            If ObjIndex >= ObjQty Then ObjIndex = 0
            Add2Selection -1
            Add2Selection ObjIndex
            ReDraw
    End Select

    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub DrawControl_KeyPress(KeyAscii As Integer)
    SuperDebug "sub/fun: DrawControl_KeyPress"
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub DrawControl_KeyUp(KeyCode As Integer, Shift As Integer)

    SuperDebug "sub/fun: DrawControl_KeyUp"
    If KeyCode >= 37 And KeyCode <= 40 And ObjIndex > -1 And Shift = 0 Then Add2UndoBuffer
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub DrawControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    On Error Resume Next
    SuperDebug "sub/fun: DrawControl_MouseDown"
    
    Dim n       As Long
    Dim tSelect As Boolean
    Dim tC      As myCoorType
    Dim minX    As Single
    Dim maxX    As Single
    Dim minY    As Single
    Dim maxY    As Single

    If NextLine = True Then
        Exit Sub
    End If

    If NewText = True And myText.Visible = True Then
        DrawControl.MousePointer = 0
        NewText = False
        DrawControl.Font = ObjList(ObjIndex).mFontName
        DrawControl.FontSize = ObjList(ObjIndex).mFontSize * mZF
        DrawControl.FontBold = ObjList(ObjIndex).mFontBold
        DrawControl.FontItalic = ObjList(ObjIndex).mFontItalic
        DrawControl.FontUnderline = ObjList(ObjIndex).mFontUnderline
        DrawControl.FontStrikethru = ObjList(ObjIndex).mFontStrikethru
        ObjList(ObjIndex).mText = myText.Text
        ObjList(ObjIndex).mWidth = myText.Width + 10 'DrawControl.TextWidth(myText.Text)
        ObjList(ObjIndex).mHeight = myText.Height 'DrawControl.TextHeight(myText.Text)

        If Trim(Len(myText.Text)) > 0 Then
            myText.Visible = False
        Else
            NewText = False
            ObjQty = ObjQty - 1
            ReDim Preserve ObjList(ObjQty)
            ReDraw
            DrawControl_MouseDown 1, 0, -5, -5
        End If

        myText.Visible = False
        NewObj = False
        'Exit Sub
    End If

    If NewObj = True Then 'set new object position
        isDown = True
        ObjList(ObjIndex).mTop = Y
        ObjList(ObjIndex).mLeft = X
    Else
        onObject = False

        toSize = checkSelection(X, Y) 'check which resize dot is clicked

        If toSize = -1 Then
            ObjIndex = -1
            ReDraw
        Else
            isResize = True
            Exit Sub
        End If

        LeftSel = 0 ' used to correct position when moving object
        TopSel = 0  '

        For n = ObjQty - 1 To 0 Step -1
            tC = GetSelPosition(ObjList(n).mLeft * mZF, ObjList(n).mTop * mZF, ObjList(n).mWidth * mZF, ObjList(n).mHeight * mZF, ObjList(n).mAngle)

            With tC
                minX = .posX1 - ObjList(n).mBorderWidth
                minY = .posY1 - ObjList(n).mBorderWidth
                maxX = .posX3 + ObjList(n).mBorderWidth
                maxY = .posY3 + ObjList(n).mBorderWidth
            End With

            If X > minX And X < maxX And Y > minY And Y < maxY Then
                tSelect = True
                LeftSel = X - ObjList(n).mLeft
                TopSel = Y - ObjList(n).mTop
            Else
                tSelect = False
            End If
   
            If tSelect = True Then
                onObject = True
                ObjIndex = n

                If Shift = 0 Then Add2Selection -1
                Add2Selection ObjIndex
                ShowSelection
                Exit For
            End If

        Next n

    End If

    DownX = X
    DownY = Y

    If ObjIndex = -1 And NewText = False Then
        QtySel = 0
        RaiseEvent ObjSelected(-1, oNone, -1, -1, -1, -1, -1, 0, -1, 0, -1, -1, -1, -1, -1, False, False, False, False, -1, -1, -1)
        Add2Selection -1
        ReDraw
    End If

    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub DrawControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
                                  
                                  If lLastMouseX = X And lLastMouseY = Y Then Exit Sub
lLastMouseX = X
lLastMouseY = Y
    SuperDebug "sub/fun: DrawControl_MouseMove"
    On Error Resume Next
    Dim tAspect As Single
    Dim n       As Long
    Dim tmp     As Single
    Dim tx      As Single
    Dim ty      As Single
    Dim tX2     As Single
    Dim tY2     As Single
    Dim tRatio  As Double
    Dim tIndex  As Long
    Dim tGr     As Integer

    If isDown = True Or NextLine = True Then ' show dot line for new object
        ReDraw
        DrawControl.DrawStyle = 2
        DrawControl.DrawWidth = 1
        DrawControl.DrawMode = 7

        Select Case ObjList(ObjIndex).mObjectType

            Case mline
                DrawControl.Line (ObjList(ObjIndex).mLeft, ObjList(ObjIndex).mTop)-(X, Y), vbYellow

            Case mArc
                tx = X
                ty = Y

                If Shift = 2 Then
                    tAspect = 1
                    ty = tx - ObjList(ObjIndex).mLeft + ObjList(ObjIndex).mTop
                    ObjList(ObjIndex).mHeight = ty - ObjList(ObjIndex).mTop
                End If

                DrawControl.ForeColor = vbYellow
                DrawArc ObjIndex, ObjList(ObjIndex).mLeft, ObjList(ObjIndex).mTop, (tx - ObjList(ObjIndex).mLeft), (ty - ObjList(ObjIndex).mTop), ObjList(ObjIndex).mPosX0 * mZF, ObjList(ObjIndex).mPosY0 * mZF, ObjList(ObjIndex).mPosX1 * mZF, ObjList(ObjIndex).mPosY1 * mZF, ObjList(ObjIndex).mPosX2 * mZF, ObjList(ObjIndex).mPosY2 * mZF, ObjList(ObjIndex).mPosX3 * mZF, ObjList(ObjIndex).mPosY3 * mZF

            Case mRectangle
                tx = X
                ty = Y

                If Shift = 2 Then
                    tAspect = 1
                    ty = tx - ObjList(ObjIndex).mLeft + ObjList(ObjIndex).mTop
                    ObjList(ObjIndex).mHeight = ty - ObjList(ObjIndex).mTop
                End If

                DrawControl.ForeColor = vbYellow
                DrawRectangle ObjList(ObjIndex).mLeft, ObjList(ObjIndex).mTop, (tx - ObjList(ObjIndex).mLeft), (ty - ObjList(ObjIndex).mTop), ObjList(ObjIndex).mAngle
        
            Case mRoundRectangle
                tx = X
                ty = Y

                If Shift = 2 Then
                    tAspect = 1
                    ty = tx - ObjList(ObjIndex).mLeft + ObjList(ObjIndex).mTop
                    ObjList(ObjIndex).mHeight = ty - ObjList(ObjIndex).mTop
                End If

                DrawControl.ForeColor = vbYellow
                DrawRoundRectangle ObjList(ObjIndex).mLeft, ObjList(ObjIndex).mTop, (tx - ObjList(ObjIndex).mLeft), (ty - ObjList(ObjIndex).mTop), ObjList(ObjIndex).mPointQty, ObjList(ObjIndex).mAngle
    
            Case mEllipse
                tx = X
                ty = Y

                If Shift = 2 Then
                    tAspect = 1
                    ty = tx - ObjList(ObjIndex).mLeft + ObjList(ObjIndex).mTop
                    ObjList(ObjIndex).mHeight = ty - ObjList(ObjIndex).mTop
                End If

                DrawControl.ForeColor = vbYellow
                DrawEllipse ObjList(ObjIndex).mLeft, ObjList(ObjIndex).mTop, (tx - ObjList(ObjIndex).mLeft), (ty - ObjList(ObjIndex).mTop), ObjList(ObjIndex).mAngle ', False

            Case mText
                tx = X
                ty = Y

                If Shift = 2 Then
                    tAspect = 1
                    ty = tx - ObjList(ObjIndex).mLeft + ObjList(ObjIndex).mTop
                    ObjList(ObjIndex).mHeight = ty - ObjList(ObjIndex).mTop
                End If

                DrawControl.ForeColor = vbYellow
                DrawRectangle ObjList(ObjIndex).mLeft, ObjList(ObjIndex).mTop, (tx - ObjList(ObjIndex).mLeft), (ty - ObjList(ObjIndex).mTop), ObjList(ObjIndex).mAngle

            Case mImage
                tx = X
                ty = Y

                If Shift = 2 Then
                    tAspect = 1
                    ty = tx - ObjList(ObjIndex).mLeft + ObjList(ObjIndex).mTop
                    tRatio = ObjList(ObjIndex).mPicture.Height / ObjList(ObjIndex).mPicture.Width
                    ObjList(ObjIndex).mWidth = tx - ObjList(ObjIndex).mLeft
                    ObjList(ObjIndex).mHeight = tRatio * ObjList(ObjIndex).mWidth
                Else
                    ObjList(ObjIndex).mWidth = tx - ObjList(ObjIndex).mLeft
                    ObjList(ObjIndex).mHeight = ty - ObjList(ObjIndex).mTop
                End If

                DrawControl.ForeColor = vbYellow
                DrawRectangle ObjList(ObjIndex).mLeft, ObjList(ObjIndex).mTop, (tx - ObjList(ObjIndex).mLeft), (ty - ObjList(ObjIndex).mTop), ObjList(ObjIndex).mAngle

                'DrawGIS2 CInt(.mWidth), CInt(.mHeight), .mPicture
            Case mPolygon
                tx = X
                ty = Y

                If Shift = 2 Then
                    tAspect = 1
                    ty = tx - ObjList(ObjIndex).mLeft + ObjList(ObjIndex).mTop
                    ObjList(ObjIndex).mHeight = ty - ObjList(ObjIndex).mTop
                End If

                DrawControl.ForeColor = vbYellow
                DrawPolygon ObjList(ObjIndex).mPointQty, ObjList(ObjIndex).mLeft, ObjList(ObjIndex).mTop, (tx - ObjList(ObjIndex).mLeft), (ty - ObjList(ObjIndex).mTop), ObjList(ObjIndex).mAngle

            Case mStar
                tx = X
                ty = Y

                If Shift = 2 Then
                    tAspect = 1
                    ty = tx - ObjList(ObjIndex).mLeft + ObjList(ObjIndex).mTop
                    ObjList(ObjIndex).mHeight = ty - ObjList(ObjIndex).mTop
                End If

                DrawControl.ForeColor = vbYellow
                DrawStar ObjList(ObjIndex).mPointQty, ObjList(ObjIndex).mLeft, ObjList(ObjIndex).mTop, (tx - ObjList(ObjIndex).mLeft), (ty - ObjList(ObjIndex).mTop), ObjList(ObjIndex).mAngle
        End Select

        DrawControl.DrawMode = 13

        RaiseEvent ObjectResize(ObjList(ObjIndex).mObjectType, ObjIndex, ObjList(ObjIndex).mLeft, ObjList(ObjIndex).mTop, tx - ObjList(ObjIndex).mLeft, ty - ObjList(ObjIndex).mTop, tAspect)

        DrawControl.DrawStyle = 0
    ElseIf Button = 1 And isDown = False Then 'resize object

        If isResize = True Then
            tRatio = ObjList(ObjIndex).mHeight / ObjList(ObjIndex).mWidth
            tx = X / mZF
            ty = Y / mZF

            Select Case toSize

                Case 0
                    tmp = ObjList(ObjIndex).mTop + ObjList(ObjIndex).mHeight
                    ObjList(ObjIndex).mTop = ty
                    ObjList(ObjIndex).mHeight = tmp - ty
                    tmp = ObjList(ObjIndex).mLeft + ObjList(ObjIndex).mWidth
                    ObjList(ObjIndex).mLeft = tx
                    ObjList(ObjIndex).mWidth = tmp - tx

                    If ObjList(ObjIndex).mObjectType = mArc Then
                        tmp = ObjList(ObjIndex).mPosY1 - ty
                        EditArc ObjIndex, toHeightN, tmp
                        tmp = ObjList(ObjIndex).mPosX0 - tx
                        EditArc ObjIndex, toWidthN, tmp
                    End If

                Case 1
                    tmp = ObjList(ObjIndex).mTop + ObjList(ObjIndex).mHeight
                    ObjList(ObjIndex).mTop = ty
                    ObjList(ObjIndex).mHeight = tmp - ty

                    If ObjList(ObjIndex).mObjectType = mArc Then
                        tmp = ObjList(ObjIndex).mPosY1 - ty
                        EditArc ObjIndex, toHeightN, tmp
                    End If

                Case 2
                    tmp = ObjList(ObjIndex).mTop + ObjList(ObjIndex).mHeight
                    ObjList(ObjIndex).mTop = ty
                    ObjList(ObjIndex).mHeight = tmp - ty
                    ObjList(ObjIndex).mWidth = tx - ObjList(ObjIndex).mLeft

                    If ObjList(ObjIndex).mObjectType = mArc Then
                        tmp = ObjList(ObjIndex).mPosY1 - ty
                        EditArc ObjIndex, toHeightN, tmp
                        tmp = tx - ObjList(ObjIndex).mPosX3
                        EditArc ObjIndex, toWidthP, tmp
                    End If

                Case 3
                    tmp = ObjList(ObjIndex).mLeft + ObjList(ObjIndex).mWidth
                    ObjList(ObjIndex).mLeft = tx
                    ObjList(ObjIndex).mWidth = tmp - tx

                    If ObjList(ObjIndex).mObjectType = mArc Then
                        tmp = ObjList(ObjIndex).mPosX0 - tx
                        EditArc ObjIndex, toWidthN, tmp
                    End If

                Case 4
                    ObjList(ObjIndex).mWidth = tx - ObjList(ObjIndex).mLeft

                    If ObjList(ObjIndex).mObjectType = mArc Then
                        tmp = tx - ObjList(ObjIndex).mPosX3
                        EditArc ObjIndex, toWidthP, tmp
                    End If

                Case 5
                    tmp = ObjList(ObjIndex).mLeft + ObjList(ObjIndex).mWidth
                    ObjList(ObjIndex).mLeft = tx
                    ObjList(ObjIndex).mWidth = tmp - tx
                    ObjList(ObjIndex).mHeight = ty - ObjList(ObjIndex).mTop

                    If ObjList(ObjIndex).mObjectType = mArc Then
                        tmp = ObjList(ObjIndex).mPosX0 - tx
                        EditArc ObjIndex, toWidthN, tmp
                        tmp = ty - ObjList(ObjIndex).mPosY0
                        EditArc ObjIndex, toHeightP, tmp
                    End If

                Case 6
                    ObjList(ObjIndex).mHeight = ty - ObjList(ObjIndex).mTop

                    If ObjList(ObjIndex).mObjectType = mArc Then
                        tmp = ty - ObjList(ObjIndex).mPosY0
                        EditArc ObjIndex, toHeightP, tmp
                    End If

                Case 7
                    ObjList(ObjIndex).mWidth = tx - ObjList(ObjIndex).mLeft
                    ObjList(ObjIndex).mHeight = ty - ObjList(ObjIndex).mTop

                    If ObjList(ObjIndex).mObjectType = mArc Then
                        tmp = tx - ObjList(ObjIndex).mPosX3
                        EditArc ObjIndex, toWidthP, tmp
                        tmp = ty - ObjList(ObjIndex).mPosY0
                        EditArc ObjIndex, toHeightP, tmp
                    End If

            End Select

            If Shift = 2 Then ObjList(ObjIndex).mHeight = tRatio * ObjList(ObjIndex).mWidth
            ReDraw
            Exit Sub
        ElseIf ObjIndex = -1 Then ' draw dot rect for mouse selection
            ReDraw
            DrawControl.DrawStyle = 2
            DrawControl.DrawMode = 7
            DrawControl.Line (DownX, DownY)-(X, Y), &H55F5F, B
            DrawControl.DrawStyle = 0
            DrawControl.DrawMode = 13
            MouseSel = True
        End If

        If onObject = True Then 'move object
            ReDraw
            DrawControl.MousePointer = 15
            DrawControl.DrawStyle = 4
            DrawControl.DrawMode = 7
            DrawControl.ForeColor = &H808080
            tx = (X - LeftSel) * mZF
            ty = (Y - TopSel) * mZF
            Xmove = 0
            Ymove = 0
            tGr = ObjList(ObjIndex).mGroupMember
    
            If QtySel > 0 And tGr = 0 Then
                Xmove = tx - ObjList(ObjIndex).mLeft
                Ymove = ty - ObjList(ObjIndex).mTop

                For n = 0 To QtySel - 1
                    tIndex = ListSel(n)
                    tX2 = ObjList(tIndex).mLeft + Xmove
                    tY2 = ObjList(tIndex).mTop + Ymove

                    Select Case ObjList(tIndex).mObjectType

                        Case mline
                            DrawControl.Line (tX2, tY2)-(tX2 + ObjList(tIndex).mWidth * mZF, tY2 + ObjList(tIndex).mHeight * mZF), &H808080
            
                        Case mArc
                            DrawArc tIndex, tX2, tY2, ObjList(tIndex).mWidth, ObjList(tIndex).mHeight, (ObjList(tIndex).mPosX0 + Xmove) * mZF, (ObjList(tIndex).mPosY0 + Ymove) * mZF, (ObjList(tIndex).mPosX1 + Xmove) * mZF, (ObjList(tIndex).mPosY1 + Ymove) * mZF, (ObjList(tIndex).mPosX2 + Xmove) * mZF, (ObjList(tIndex).mPosY2 + Ymove) * mZF, (ObjList(tIndex).mPosX3 + Xmove) * mZF, (ObjList(tIndex).mPosY3 + Ymove) * mZF
             
                        Case mRectangle
                            DrawRectangle tX2, tY2, ObjList(tIndex).mWidth * mZF, ObjList(tIndex).mHeight * mZF, ObjList(tIndex).mAngle
            
                        Case mRoundRectangle
                            DrawRoundRectangle tX2, tY2, ObjList(tIndex).mWidth * mZF, ObjList(tIndex).mHeight * mZF, ObjList(tIndex).mPointQty, ObjList(tIndex).mAngle
            
                        Case mEllipse
                            DrawEllipse tX2, tY2, ObjList(tIndex).mWidth * mZF, ObjList(tIndex).mHeight * mZF, ObjList(tIndex).mAngle ', False
            
                        Case mText
                            DrawRectangle tX2, tY2, ObjList(tIndex).mWidth * mZF, ObjList(tIndex).mHeight * mZF, ObjList(tIndex).mAngle

                        Case mImage
                            DrawRectangle tX2, tY2, ObjList(tIndex).mWidth * mZF, ObjList(tIndex).mHeight * mZF, ObjList(tIndex).mAngle
         
                        Case mPolygon
                            DrawPolygon ObjList(tIndex).mPointQty, tX2, tY2, ObjList(tIndex).mWidth * mZF, ObjList(tIndex).mHeight * mZF, ObjList(tIndex).mAngle
       
                        Case mStar
                            DrawStar ObjList(tIndex).mPointQty, tX2, tY2, ObjList(tIndex).mWidth * mZF, ObjList(tIndex).mHeight * mZF, ObjList(tIndex).mAngle
        
                    End Select

                Next n
        
            ElseIf tGr > 0 Then
                Xmove = tx - ObjList(ObjIndex).mLeft
                Ymove = ty - ObjList(ObjIndex).mTop

                For n = 0 To ObjQty - 1

                    If ObjList(n).mGroupMember = tGr Then
                        tX2 = ObjList(n).mLeft + Xmove
                        tY2 = ObjList(n).mTop + Ymove

                        Select Case ObjList(n).mObjectType

                            Case mline
                                DrawControl.Line (tX2, tY2)-(tX2 + ObjList(n).mWidth * mZF, tY2 + ObjList(n).mHeight * mZF), &H808080
                
                            Case mArc
                                DrawArc n, tX2, tY2, ObjList(n).mWidth, ObjList(n).mHeight, (ObjList(n).mPosX0 + Xmove) * mZF, (ObjList(n).mPosY0 + Ymove) * mZF, (ObjList(n).mPosX1 + Xmove) * mZF, (ObjList(n).mPosY1 + Ymove) * mZF, (ObjList(n).mPosX2 + Xmove) * mZF, (ObjList(n).mPosY2 + Ymove) * mZF, (ObjList(n).mPosX3 + Xmove) * mZF, (ObjList(n).mPosY3 + Ymove) * mZF
                 
                            Case mRectangle
                                DrawRectangle tX2, tY2, ObjList(n).mWidth * mZF, ObjList(n).mHeight * mZF, ObjList(n).mAngle
                
                            Case mRoundRectangle
                                DrawRoundRectangle tX2, tY2, ObjList(n).mWidth * mZF, ObjList(n).mHeight * mZF, ObjList(n).mPointQty, ObjList(n).mAngle
                
                            Case mEllipse
                                DrawEllipse tX2, tY2, ObjList(n).mWidth * mZF, ObjList(n).mHeight * mZF, ObjList(n).mAngle ', False
                
                            Case mText
                                DrawRectangle tX2, tY2, ObjList(n).mWidth * mZF, ObjList(n).mHeight * mZF, ObjList(n).mAngle
    
                            Case mImage
                                DrawControl.Line (tX2, tY2)-(tX2 + ObjList(n).mWidth * mZF, tY2 + ObjList(n).mHeight * mZF), &H808080, B
                
                            Case mPolygon
                                DrawPolygon ObjList(n).mPointQty, tX2, tY2, ObjList(n).mWidth * mZF, ObjList(n).mHeight * mZF, ObjList(n).mAngle
           
                            Case mStar
                                DrawStar ObjList(n).mPointQty, tX2, tY2, ObjList(n).mWidth * mZF, ObjList(n).mHeight * mZF, ObjList(n).mAngle
            
                        End Select

                    End If

                Next n

            End If

            DrawControl.DrawMode = 13
            DrawControl.DrawStyle = 0
            isMove = True

            If NewText = False Then
                
                RaiseEvent ObjSelected(ObjList(ObjIndex).mObjectType, ObjList(ObjIndex).mGISObj, ObjIndex, tx, ty, ObjList(ObjIndex).mWidth, ObjList(ObjIndex).mHeight, ObjList(ObjIndex).mAngle, ObjList(ObjIndex).mFillColor, ObjList(ObjIndex).mFillStyle, ObjList(ObjIndex).mBorderColor, ObjList(ObjIndex).mBorderWidth, ObjList(ObjIndex).mAspect, ObjList(ObjIndex).mFontName, ObjList(ObjIndex).mFontSize, ObjList(ObjIndex).mFontBold, ObjList(ObjIndex).mFontItalic, ObjList(ObjIndex).mFontUnderline, ObjList(ObjIndex).mFontStrikethru, ObjList(ObjIndex).mText, ObjList(ObjIndex).mTextAlign, ObjList(ObjIndex).mPointQty)
            End If
        End If
    End If

    RaiseEvent MouseMove(Button, Shift, X / mZF, Y / mZF)
End Sub

Private Sub DrawControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    SuperDebug "sub/fun: DrawControl_MouseUp"
    On Error Resume Next
    Dim tBorderColor As Long
    Dim tWidth       As Integer
    Dim tIndex       As Long
    Dim n            As Long
    Dim tC           As myCoorType
    Dim tGr          As Integer

    DrawControl.MousePointer = 0

    If NextLine = True Then
        If Button = 2 Then
            NextLine = False

            NewObj = False
            ObjQty = ObjQty - 1
            ReDim Preserve ObjList(ObjQty)
            ReDraw
            DrawControl_MouseDown 1, 0, -5, -5
        End If
    End If

    If isResize = True Then
        Add2UndoBuffer
        isResize = False
        ReDraw
    End If

    If NewObj = True Then
        NewObj = False
        isDown = False

        If ObjList(ObjIndex).mObjectType <> mline Then

            With ObjList(ObjIndex)
                .mLeft = .mLeft / mZF
                .mTop = .mTop / mZF
                .mHeight = (Y / mZF - .mTop)
                .mWidth = (X / mZF - .mLeft)

                If Shift = 2 Then .mAspect = 1 Else .mAspect = 0
                tBorderColor = .mBorderColor
                tWidth = .mBorderWidth

                If ObjList(ObjIndex).mObjectType = mArc Then InitArc ObjIndex
                ReDraw
            End With

        ElseIf ObjList(ObjIndex).mObjectType = mline Then
            ObjList(ObjIndex).mHeight = (Y - ObjList(ObjIndex).mTop) / mZF
            ObjList(ObjIndex).mWidth = (X - ObjList(ObjIndex).mLeft) / mZF
            tBorderColor = ObjList(ObjIndex).mBorderColor
            tWidth = ObjList(ObjIndex).mBorderWidth
            AddObject mline, Y, X, , , , , , tBorderColor, tWidth
            NewObj = True
            NextLine = True

            DrawControl.MousePointer = 99
            Set DrawControl.MouseIcon = cLine.Picture
        End If

        If NewText = True And myText.Visible = False Then

            With ObjList(ObjIndex)
                DrawControl.MousePointer = 3
                myText.left = .mLeft
                myText.top = .mTop
                myText.Font = .mFontName
                myText.FontSize = .mFontSize * mZF
                myText.FontBold = .mFontBold
                myText.FontItalic = .mFontItalic
                myText.FontUnderline = .mFontUnderline
                myText.FontStrikethru = .mFontStrikethru
                myText.Alignment = .mTextAlign
                DrawControl.FontName = .mFontName
                DrawControl.FontSize = .mFontSize
                DrawControl.FontBold = .mFontBold
                DrawControl.FontItalic = .mFontItalic
                DrawControl.FontUnderline = .mFontUnderline
                DrawControl.FontStrikethru = .mFontStrikethru

                If Len(.mText) = 0 Then
                    myText.Text = mDefaultText
                Else
                    myText.Text = .mText
                End If

                If .mWidth < DrawControl.TextWidth(myText.Text) Then
                    myText.Width = DrawControl.TextWidth(myText.Text)
                    .mWidth = myText.Width
                Else
                    myText.Width = .mWidth
                End If

                If .mHeight < DrawControl.TextHeight(myText.Text) Then
                    myText.Height = DrawControl.TextHeight(myText.Text)
                    .mHeight = myText.Height
                Else
                    myText.Height = .mHeight
                End If

                myText.Visible = True
                myText.SelStart = 0
                myText.SelLength = Len(myText.Text)
                myText.SetFocus
            End With

        End If

        Add2UndoBuffer
        RaiseEvent NewDrawingEnd
    ElseIf Button = 1 And onObject = True And isMove = True Then
        isMove = False
        tGr = ObjList(ObjIndex).mGroupMember

        If QtySel > 0 And tGr = 0 Then

            For n = 0 To QtySel - 1
                tIndex = ListSel(n)

                With ObjList(tIndex)
                    .mLeft = (.mLeft + Xmove) / mZF
                    .mTop = (.mTop + Ymove) / mZF

                    If .mObjectType = mArc Then
                        .mPosX0 = (.mPosX0 + Xmove)
                        .mPosY0 = (.mPosY0 + Ymove)
                        .mPosX1 = (.mPosX1 + Xmove)
                        .mPosY1 = (.mPosY1 + Ymove)
                        .mPosX2 = (.mPosX2 + Xmove)
                        .mPosY2 = (.mPosY2 + Ymove)
                        .mPosX3 = (.mPosX3 + Xmove)
                        .mPosY3 = (.mPosY3 + Ymove)
                    End If

                End With

            Next n

        ElseIf tGr > 0 Then

            For n = 0 To ObjQty - 1

                With ObjList(n)

                    If .mGroupMember = tGr Then
                        .mLeft = (.mLeft + Xmove) / mZF
                        .mTop = (.mTop + Ymove) / mZF
                    End If

                End With

            Next n

        End If

        Add2UndoBuffer
        ReDraw
    ElseIf Button = 1 And MouseSel = True Then

        For n = ObjQty - 1 To 0 Step -1
            tC = GetSelPosition(ObjList(n).mLeft, ObjList(n).mTop, ObjList(n).mWidth, ObjList(n).mHeight, ObjList(n).mAngle)

            With tC

                If .posX1 > DownX And .posY1 > DownY And .posX3 < X And .posY3 < Y Or .posX1 > X And .posY1 > Y And .posX3 < DownX And .posY3 < DownY Or .posX1 > X And .posY1 > DownY And .posX3 < DownX And .posY3 < Y Or .posX1 > DownX And .posY1 > Y And .posX3 < X And .posY3 < DownY Then

                    If QtySel = 0 Then ObjIndex = n
                    Add2Selection n
                End If

            End With

        Next n

        ShowSelection
        MouseSel = False
        ReDraw
    End If

    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub ReDraw(Optional ShowSel As Boolean = True)
    SuperDebug "sub/fun: ReDraw"
    On Error Resume Next
    Dim tRatio As Double
    Dim n      As Long

    DrawControl.Cls
    DrawControl.DrawMode = 13

    If ObjQty > 0 Then

        For n = 0 To ObjQty - 1

            With ObjList(n)

                If .mDeleted = False Then

                    DrawControl.FillStyle = .mFillStyle
    
                    If .mBorderWidth <= 0 Then
                        DrawControl.DrawStyle = 5
                    Else
                        DrawControl.DrawStyle = 0
                    End If
    
                    Select Case .mObjectType

                        Case mline

                            If DrawControl.DrawWidth < 1 Then DrawControl.DrawWidth = 1
                            DrawControl.DrawWidth = .mBorderWidth * mZF
                            DrawControl.ForeColor = .mBorderColor
                            DrawControl.Line (.mLeft * mZF, .mTop * mZF)-((.mLeft * mZF) + (.mWidth * mZF), (.mTop * mZF + .mHeight * mZF)), Abs(.mBorderColor)
            
                        Case mArc
                            DrawControl.FillColor = .mFillColor
                            DrawControl.DrawWidth = .mBorderWidth * mZF
                            DrawControl.ForeColor = .mBorderColor
                            DrawArc n, .mLeft * mZF, .mTop * mZF, .mWidth * mZF, .mHeight * mZF, .mPosX0 * mZF, .mPosY0 * mZF, .mPosX1 * mZF, .mPosY1 * mZF, .mPosX2 * mZF, .mPosY2 * mZF, .mPosX3 * mZF, .mPosY3 * mZF

                        Case mRectangle

                            If .mAspect = 1 Then
                                .mHeight = .mWidth
                            End If

                            DrawControl.DrawWidth = .mBorderWidth * mZF
                            DrawControl.FillColor = .mFillColor
                            DrawControl.ForeColor = .mBorderColor
                            DrawRectangle .mLeft * mZF, .mTop * mZF, .mWidth * mZF, .mHeight * mZF, .mAngle
            
                        Case mRoundRectangle

                            If .mAspect = 1 Then
                                .mHeight = .mWidth
                            End If

                            DrawControl.DrawWidth = .mBorderWidth * mZF
                            DrawControl.FillColor = .mFillColor
                            DrawControl.ForeColor = .mBorderColor
                            DrawRoundRectangle .mLeft * mZF, .mTop * mZF, .mWidth * mZF, .mHeight * mZF, .mPointQty, .mAngle
            
                        Case mEllipse

                            If .mAspect = 1 Then
                                .mHeight = .mWidth
                            End If

                            DrawControl.FillColor = .mFillColor
                            DrawControl.DrawWidth = .mBorderWidth * mZF
                            DrawControl.ForeColor = .mBorderColor
                            DrawEllipse .mLeft * mZF, .mTop * mZF, .mWidth * mZF, .mHeight * mZF, .mAngle

                        Case mText

                            If Len(.mText) > 0 And NewText = False Then
                                DrawControl.FillColor = .mFillColor
                                DrawControl.ForeColor = .mFillColor
                                DrawText .mText, .mLeft * mZF, .mTop * mZF, .mWidth * mZF, .mHeight * mZF, .mFontName, .mFontSize * mZF, .mAngle, .mFontBold, .mFontItalic, .mFontUnderline, .mFontStrikethru, .mTextAlign
                            End If

                        Case mImage

                            If .mAspect = 1 Then
                                tRatio = .mPicture.Height / .mPicture.Width
                                .mHeight = .mWidth * tRatio
                            End If

                            If .mWidth = -1 Then
                                .mWidth = DrawControl.ScaleX(.mPicture.Width)
                            End If

                            If .mHeight = -1 Then
                                .mHeight = DrawControl.ScaleY(.mPicture.Height)
                            End If

                            If .mGISObj = oMap Then
                                'TODO Petri
                                DrawGIS2 CInt(.mLeft), CInt(.mTop), CInt(.mWidth), CInt(.mHeight), .mPicture
                            ElseIf .mGISObj = oLegend Then
                                DrawLegend CInt(.mLeft), CInt(.mTop), CInt(.mWidth), CInt(.mHeight), .mPicture
                            ElseIf .mGISObj = oNortArrow Then
                                DrawNorthArrow CInt(.mLeft), CInt(.mTop), CInt(.mWidth), CInt(.mHeight), .mPicture
                            ElseIf .mGISObj = oScaleBar Then
                                DrawScaleBar CInt(.mLeft), CInt(.mTop), CInt(.mWidth), CInt(.mHeight), .mPicture
                            ElseIf .mGISObj = oDataGrid Then
                            
                            End If
                            
                            DrawPicture .mPicture, .mLeft * mZF, .mTop * mZF, .mWidth * mZF, .mHeight * mZF, .mAngle
            
                        Case mPolygon

                            If .mAspect = 1 Then
                                .mHeight = .mWidth
                            End If

                            DrawControl.DrawWidth = .mBorderWidth * mZF
                            DrawControl.FillColor = .mFillColor
                            DrawControl.ForeColor = .mBorderColor
                            DrawPolygon .mPointQty, .mLeft * mZF, .mTop * mZF, .mWidth * mZF, .mHeight * mZF, .mAngle
            
                        Case mStar

                            If .mAspect = 1 Then
                                .mHeight = .mWidth
                            End If

                            DrawControl.DrawWidth = .mBorderWidth * mZF
                            DrawControl.FillColor = .mFillColor
                            DrawControl.ForeColor = .mBorderColor
                            DrawStar .mPointQty, .mLeft * mZF, .mTop * mZF, .mWidth * mZF, .mHeight * mZF, .mAngle
                    End Select
    
                End If

            End With

        Next n

    End If

    DrawControl.FillStyle = 1
    DrawControl.DrawWidth = 1
    DrawControl.DrawStyle = 0
    DrawControl.Font = myFont
    DrawControl.FontSize = 8
    DrawControl.FontBold = False
    DrawControl.FontItalic = False
    DrawControl.FontUnderline = False
    DrawControl.FontStrikethru = False

    If isDown = False And NextLine = False And NewText = False Then
        If ShowSel = True Then
            ShowSelection
        End If
    End If

End Sub

Private Function checkSelection(selX As Single, selY As Single) As Integer
    SuperDebug "sub/fun: checkSelection"
    On Error Resume Next            ' check which selection dot is clicked
    Dim tC As myCoorType
    Dim TS As Integer

    If ObjIndex = -1 Then
        checkSelection = -1
        DrawControl.MousePointer = 0
        Exit Function
    End If

    If ObjList(ObjIndex).mGroupMember > 0 Then 'to avoid resize on grouped objects
        checkSelection = -1
        DrawControl.MousePointer = 0
        Exit Function
    End If

    tC = GetSelPosition(ObjList(ObjIndex).mLeft * mZF, ObjList(ObjIndex).mTop * mZF, ObjList(ObjIndex).mWidth * mZF, ObjList(ObjIndex).mHeight * mZF, ObjList(ObjIndex).mAngle)

    If selX > tC.posX1 - 10 And selY > tC.posY1 - 10 And selX < tC.posX1 - 2 And selY < tC.posY1 - 2 Then
        TS = 0
        DrawControl.MousePointer = 8
    ElseIf selX > tC.posX1 + tC.centerX - 4 And tC.posY1 - 10 And selX < tC.posX1 + tC.centerX + 4 And selY < tC.posY1 - 2 Then
        TS = 1
        DrawControl.MousePointer = 7
    ElseIf selX > tC.posX4 + 2 And selY > tC.posY4 - 10 And selX < tC.posX4 + 10 And selY < tC.posY4 - 2 Then
        TS = 2
        DrawControl.MousePointer = 6
    ElseIf selX > tC.posX1 - 10 And selY > tC.posY1 + tC.centerY - 4 And selX < tC.posX1 - 2 And selY < tC.posY1 + tC.centerY + 4 Then
        TS = 3
        DrawControl.MousePointer = 9
    ElseIf selX > tC.posX4 + 2 And selY > ((tC.posY4 - tC.posY3) / 2) + tC.posY3 - 4 And selX < ((tC.posX4 - tC.posX3) / 2) + tC.posX3 + 10 And selY < ((tC.posY4 - tC.posY3) / 2) + tC.posY3 + 4 Then
        TS = 4
        DrawControl.MousePointer = 9
    ElseIf selX > tC.posX2 - 10 And selY > tC.posY2 + 2 And selX < tC.posX2 - 2 And selY < tC.posY2 + 10 Then
        TS = 5
        DrawControl.MousePointer = 6
    ElseIf selX > tC.posX1 + tC.centerX - 4 And selY > tC.posY2 + 2 And selX < tC.posX1 + tC.centerX + 4 And selY < tC.posY2 + 10 Then
        TS = 6
        DrawControl.MousePointer = 7
    ElseIf selX > tC.posX3 + 2 And selY > tC.posY3 + 2 And selX < tC.posX3 + 10 And selY < tC.posY3 + 10 Then
        TS = 7
        DrawControl.MousePointer = 8
    Else
        TS = -1
    End If

    If ObjList(ObjIndex).mWidth < 0 And ObjList(ObjIndex).mHeight < 0 And TS >= 0 Then

        Select Case TS

            Case 0: checkSelection = 7

            Case 1: checkSelection = 6

            Case 2: checkSelection = 5

            Case 3: checkSelection = 4

            Case 4: checkSelection = 3

            Case 5: checkSelection = 2

            Case 6: checkSelection = 1

            Case 7: checkSelection = 0
        End Select

    ElseIf ObjList(ObjIndex).mWidth < 0 And TS >= 0 Then

        Select Case TS

            Case 0: checkSelection = 2

            Case 1: checkSelection = 1

            Case 2: checkSelection = 0

            Case 3: checkSelection = 4

            Case 4: checkSelection = 3

            Case 5: checkSelection = 7

            Case 6: checkSelection = 6

            Case 7: checkSelection = 5
        End Select

    ElseIf ObjList(ObjIndex).mHeight < 0 And TS >= 0 Then

        Select Case TS

            Case 0: checkSelection = 5

            Case 1: checkSelection = 6

            Case 2: checkSelection = 7

            Case 3: checkSelection = 3

            Case 4: checkSelection = 4

            Case 5: checkSelection = 0

            Case 6: checkSelection = 1

            Case 7: checkSelection = 2
        End Select

    Else
        checkSelection = TS
    End If

End Function

Public Sub ModifyObject(Optional tTop As Single = -1, _
                        Optional tLeft As Single = -1, _
                        Optional tHeight As Single = -1, _
                        Optional tWidth As Single = -1, _
                        Optional tAngle As Single = -1, _
                        Optional tFillColor As Long = -1, _
                        Optional tFillStyle As myFill = -1, _
                        Optional tBorderColor As Long = -1, _
                        Optional tBorderWidth As Integer = -1, _
                        Optional tPicture As StdPicture, _
                        Optional tFontName As String = "", _
                        Optional tFontSize As Integer = -1, _
                        Optional tFontBold As myBool3 = Unchanged, _
                        Optional tFontItalic As myBool3 = Unchanged, _
                        Optional tFontUnderline As myBool3 = Unchanged, _
                        Optional tFontStrikethru As myBool3 = Unchanged, _
                        Optional tText As String = "", _
                        Optional tTextAlign As AlignmentConstants = -1, _
                        Optional tPointQty As Integer = -1, _
                        Optional tPosX0 As Single = -1, _
                        Optional tPosY0 As Single = -1, _
                        Optional tPosX1 As Single = -1, _
                        Optional tPosY1 As Single = -1, _
                        Optional tPosX2 As Single = -1, _
                        Optional tPosY2 As Single = -1, Optional tPosX3 As Single = -1, Optional tPosY3 As Single = -1, Optional tGroupMember As Integer = -1)

    SuperDebug "sub/fun: ModifyObject"
    Dim n   As Long
    Dim tGr As Integer

    NextLine = False

    NewObj = False

    On Error Resume Next

    tGr = ObjList(ObjIndex).mGroupMember

    If QtySel > 0 And tGr = 0 Then

        For n = 0 To QtySel - 1

            With ObjList(ListSel(n))

                If tFillColor > -1 Then .mFillColor = tFillColor
                If tFillStyle > -1 Then .mFillStyle = tFillStyle
                If tAngle > -1 Then .mAngle = tAngle
                If .mObjectType = mArc Then .mAngle = 0
                If tBorderColor > -1 Then .mBorderColor = tBorderColor
                If tBorderWidth > -1 Then .mBorderWidth = tBorderWidth
                If tPointQty > -1 Then .mPointQty = tPointQty
                If tPosX0 > -1 Then .mPosX0 = tPosX0
                If tPosY0 > -1 Then .mPosY0 = tPosY0
                If tPosX1 > -1 Then .mPosX1 = tPosX1
                If tPosY1 > -1 Then .mPosY1 = tPosY1
                If tPosX2 > -1 Then .mPosX2 = tPosX2
                If tPosY2 > -1 Then .mPosY2 = tPosY2
                If tPosX3 > -1 Then .mPosX3 = tPosX3
                If tPosY3 > -1 Then .mPosY3 = tPosY3
                If tGroupMember > -1 Then .mGroupMember = tGroupMember
                If .mObjectType = mText Then
                    If tFontName <> "" Then .mFontName = tFontName
                    If tFontSize > 2 Then .mFontSize = tFontSize
                    If tFontBold <> Unchanged Then .mFontBold = tFontBold
                    If tFontItalic <> Unchanged Then .mFontItalic = tFontItalic
                    If tFontUnderline <> Unchanged Then .mFontUnderline = tFontUnderline
                    If tFontStrikethru <> Unchanged Then .mFontStrikethru = tFontStrikethru
                    If tTextAlign <> -1 Then .mTextAlign = tTextAlign
                    DrawControl.FontName = .mFontName
                    DrawControl.FontSize = .mFontSize
                    DrawControl.FontBold = .mFontBold
                    DrawControl.FontItalic = .mFontItalic
                    DrawControl.FontUnderline = .mFontUnderline
                    DrawControl.FontStrikethru = .mFontStrikethru

                    If tAngle > -1 Then .mAngle = tAngle
                End If

            End With

        Next n

        ReDraw
        Add2UndoBuffer
    ElseIf tGr > 0 Then

        For n = 0 To ObjQty - 1

            With ObjList(n)

                If .mGroupMember = tGr Then
                    If tFillColor > -1 Then .mFillColor = tFillColor
                    If tFillStyle > -1 Then .mFillStyle = tFillStyle
                    If tAngle > -1 Then .mAngle = tAngle
                    If tBorderColor > -1 Then .mBorderColor = tBorderColor
                    If tBorderWidth > -1 Then .mBorderWidth = tBorderWidth
                    If tPointQty > -1 Then .mPointQty = tPointQty
                    If tPosX1 > -1 Then .mPosX1 = tPosX1
                    If tPosX2 > -1 Then .mPosX2 = tPosX2
                    If tPosY1 > -1 Then .mPosY1 = tPosY1
                    If tPosY2 > -1 Then .mPosY2 = tPosY2
                    If tGroupMember > -1 Then .mGroupMember = tGroupMember
                    If .mObjectType = mText Then
                        If tFontName <> "" Then .mFontName = tFontName
                        If tFontSize > 2 Then .mFontSize = tFontSize
                        If tFontBold <> Unchanged Then .mFontBold = tFontBold
                        If tFontItalic <> Unchanged Then .mFontItalic = tFontItalic
                        If tFontUnderline <> Unchanged Then .mFontUnderline = tFontUnderline
                        If tFontStrikethru <> Unchanged Then .mFontStrikethru = tFontStrikethru
                        DrawControl.FontName = .mFontName
                        DrawControl.FontSize = .mFontSize
                        DrawControl.FontBold = .mFontBold
                        DrawControl.FontItalic = .mFontItalic
                        DrawControl.FontUnderline = .mFontUnderline
                        DrawControl.FontStrikethru = .mFontStrikethru

                        If tAngle > -1 Then .mAngle = tAngle
                    End If
                End If

            End With

        Next n

        ReDraw
        Add2UndoBuffer
    End If

End Sub

Public Sub Export2BMP(Filename As String)
    SuperDebug "sub/fun: Export2BMP"
    SavePicture DrawControl.Image, Filename
End Sub

Public Sub DeleteObj(Optional tObjIndex As Long = -1)
    SuperDebug "sub/fun: DeleteObj"
    Dim n   As Long
    Dim tGr As Integer
    NextLine = False

    NewObj = False

    If tObjIndex = -1 Then
        If QtySel > 0 Then
            Add2UndoBuffer
            tGr = ObjList(ObjIndex).mGroupMember

            If tGr = 0 Then

                For n = 0 To QtySel - 1
                    ObjList(ListSel(n)).mDeleted = True
                Next n

            ElseIf tGr > 0 Then

                For n = 0 To ObjQty - 1

                    If ObjList(n).mGroupMember = tGr Then
                        ObjList(n).mDeleted = True
                    End If

                Next n

            End If
        End If

    Else
        ObjList(ListSel(tObjIndex)).mDeleted = True
    End If

    Add2Selection -1

    RedimList
    ReDraw

    RaiseEvent ObjSelected(-1, oNone, -1, -1, -1, -1, -1, 0, -1, 0, -1, -1, -1, -1, -1, False, False, False, False, -1, -1, -1)

End Sub

Private Sub RedimList()
    SuperDebug "sub/fun: RedimList"
    Dim tmpList() As myObject
    Dim n1        As Long
    Dim n2        As Long

    n2 = 0

    For n1 = 0 To ObjQty - 1

        If ObjList(n1).mDeleted = False Then
            n2 = n2 + 1
            ReDim Preserve tmpList(n2)
            tmpList(n2 - 1) = ObjList(n1)
        End If

    Next n1

    ReDim ObjList(n2)
    ObjQty = n2

    For n1 = 0 To n2 - 1
        ObjList(n1) = tmpList(n1)
    Next n1

    ObjIndex = -1
End Sub

Public Sub SaveProjects(Filename As String)
    SuperDebug "sub/fun: SaveProjects"
    Dim n       As Long
    Dim FF      As Integer
    Dim mFile   As String
    Dim mData   As String
    Dim tmpFile As String

    If Len(Dir(Filename)) Then Kill Filename

    If ObjQty > 0 Then
        mFile = "ObjDrawFile" & String(3, 0)

        For n = 0 To ObjQty - 1

            With ObjList(n)
                mData = mData & String(5, 254) & String(5, 255) & String(5, 254) & .mObjectType & Chr(0) & Hex(.mTop) & Chr(0) & Hex(.mLeft) & Chr(0) & Hex(.mHeight) & Chr(0) & Hex(.mWidth) & Chr(0) & Hex(.mAngle) & Chr(0) & Hex(DrawControl.BackColor) & Chr(0) & Hex(.mFillColor) & Chr(0) & Hex(.mFillStyle) & Chr(0) & Hex(.mBorderColor) & Chr(0) & Hex(.mBorderWidth) & Chr(0) & Hex(.mAspect) & Chr(0) & .mFontName & Chr(0) & .mFontSize & Chr(0) & Int(.mFontBold) & Chr(0) & Int(.mFontItalic) & Chr(0) & Int(.mFontUnderline) & Chr(0) & Int(.mFontStrikethru) & Chr(0) & .mText & Chr(0) & .mTextAlign & Chr(0) & Hex(.mPointQty) & Chr(0) & Hex(.mPosX0) & Chr(0) & Hex(.mPosY0) & Chr(0) & Hex(.mPosX1) & Chr(0) & Hex(.mPosY1) & Chr(0) & Hex(.mPosX2) & Chr(0) & Hex(.mPosY2) & Chr(0) & Hex(.mPosX3) & Chr(0) & Hex(.mPosY3) & Chr(0) & Hex(.mGroupMember) & Chr(0) & Hex(mCanvasWidth) & Chr(0) & Hex(mCanvasHeight) & Chr(0) & Hex(GroupQty) & Chr(0) & Hex(.mGISObj)
    
                If ObjList(n).mObjectType = mImage Then
                    Set PicData.Picture = ObjList(n).mPicture
                    SavePicture PicData, App.Path & "\Tmp.bmp"

                    DoEvents
                    FF = FreeFile
                    Open App.Path & "\Tmp.bmp" For Binary As #FF
                    tmpFile = Input(FileLen(App.Path & "\Tmp.bmp"), #FF)
                    Close #FF

                    DoEvents
                    mData = mData & String(5, 255) & String(5, 254) & String(5, 255) & tmpFile

                    If Len(Dir(App.Path & "\Tmp.bmp")) Then Kill App.Path & "\Tmp.bmp"

                    DoEvents
                End If

            End With

        Next n
    
        FF = FreeFile
        Open Filename For Binary Access Write As #FF
        Put #FF, , mFile & mData
        Close #FF

        DoEvents
    End If

End Sub

Public Sub OpenProjects(Filename As String)
    SuperDebug "sub/fun: OpenProjects"
    On Error Resume Next
    Dim n       As Long
    Dim FF      As Integer
    Dim mFile   As String
    Dim mData   As String
    Dim tmpFile As String
    Dim tmpPic  As String

    FF = FreeFile
    Open Filename For Binary As #FF
    tmpFile = Input(FileLen(Filename), #FF)
    Close #FF

    If LCase(left(tmpFile, 11)) = "objdrawfile" Then
        ObjQty = UBound(Split(tmpFile, String(5, 254) & String(5, 255) & String(5, 254)))

        ReDim ObjList(ObjQty)

        For n = 0 To ObjQty - 1
            mFile = Split(tmpFile, String(5, 254) & String(5, 255) & String(5, 254))(n + 1)
            mData = Split(mFile, String(5, 255) & String(5, 254) & String(5, 255))(0)

            With ObjList(n)
                .mObjectType = Int(Split(mData, Chr(0))(0))
                .mTop = CLng("&H" & Split(mData, Chr(0))(1))
                .mLeft = CLng("&H" & Split(mData, Chr(0))(2))
                .mHeight = CLng("&H" & Split(mData, Chr(0))(3))
                .mWidth = CLng("&H" & Split(mData, Chr(0))(4))
                .mAngle = CLng("&H" & Split(mData, Chr(0))(5))
                DrawControl.BackColor = CLng("&H" & Split(mData, Chr(0))(6))
                .mFillColor = CLng("&H" & Split(mData, Chr(0))(7))
                .mFillStyle = CLng("&H" & Split(mData, Chr(0))(8))
                .mBorderColor = CLng("&H" & Split(mData, Chr(0))(9))
                .mBorderWidth = CLng("&H" & Split(mData, Chr(0))(10))
                .mAspect = CLng("&H" & Split(mData, Chr(0))(11))
                .mFontName = Split(mData, Chr(0))(12)
                .mFontSize = Split(mData, Chr(0))(13)
                .mFontBold = CBool(Split(mData, Chr(0))(14))
                .mFontItalic = CBool(Split(mData, Chr(0))(15))
                .mFontUnderline = CBool(Split(mData, Chr(0))(16))
                .mFontStrikethru = CBool(Split(mData, Chr(0))(17))
                .mText = Split(mData, Chr(0))(18)
                .mTextAlign = Split(mData, Chr(0))(19)
                .mPointQty = CLng("&H" & Split(mData, Chr(0))(20))
                .mPosX0 = CLng("&H" & Split(mData, Chr(0))(21))
                .mPosY0 = CLng("&H" & Split(mData, Chr(0))(22))
                .mPosX1 = CLng("&H" & Split(mData, Chr(0))(23))
                .mPosY1 = CLng("&H" & Split(mData, Chr(0))(24))
                .mPosX2 = CLng("&H" & Split(mData, Chr(0))(25))
                .mPosY2 = CLng("&H" & Split(mData, Chr(0))(26))
                .mPosX3 = CLng("&H" & Split(mData, Chr(0))(27))
                .mPosY3 = CLng("&H" & Split(mData, Chr(0))(28))
                .mGroupMember = CLng("&H" & Split(mData, Chr(0))(29))
                mCanvasWidth = CLng("&H" & Split(mData, Chr(0))(30))
                mCanvasHeight = CLng("&H" & Split(mData, Chr(0))(31))
                GroupQty = CLng("&H" & Split(mData, Chr(0))(32))
                .mGISObj = CLng("&H" & Split(mData, Chr(0))(33))

                If .mObjectType = mImage Then
                    tmpPic = Split(mFile, String(5, 255) & String(5, 254) & String(5, 255))(1)
                    FF = FreeFile
                    Open App.Path & "\Tmp.bmp" For Binary Access Write As #FF
                    Put #FF, , tmpPic
                    Close #FF

                    DoEvents
                    Set .mPicture = LoadPicture(App.Path & "\Tmp.bmp")

                    DoEvents

                    If Len(Dir(App.Path & "\Tmp.bmp")) Then Kill App.Path & "\Tmp.bmp"
                End If

            End With

        Next n

    End If

    ObjIndex = -1
    mZF = 1
    UserControl_Resize
    ReDraw
    RaiseEvent UndoRedo(True, True)
    ReDim UndoBuffer(mUndoSize)
    UndoStack = 0
    UndoPointer = 0

End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = DrawControl.BackColor
End Property

Public Property Let BackColor(ByVal vNewBackColor As OLE_COLOR)
    SuperDebug "sub/fun: BackColor"
    DrawControl.BackColor = vNewBackColor
    ReDraw
    PropertyChanged "BackColor"
End Property

Public Sub SetObjectOrder(tObjIndex As Long, tOrder As myOrder)
    SuperDebug "sub/fun: SetObjectOrder"
    On Error Resume Next
    Dim tmpList() As myObject
    Dim n1        As Long
    Dim n2        As Long

    n2 = 0

    NextLine = False

    NewObj = False

    ReDim tmpList(ObjQty)

    Select Case tOrder

        Case BringToFront

            For n1 = 0 To ObjQty - 1
                n2 = n2 + 1

                If n1 < tObjIndex Then
                    tmpList(n2 - 1) = ObjList(n1)
                ElseIf n1 = tObjIndex Then
                    tmpList(ObjQty - 1) = ObjList(tObjIndex)
                ElseIf n1 > tObjIndex Then
                    tmpList(n2 - 2) = ObjList(n1)
                End If

            Next n1

            Add2Selection -1
            ObjIndex = ObjQty - 1
            Add2Selection ObjIndex

        Case SendToBack
            tmpList(0) = ObjList(tObjIndex)

            For n1 = 0 To ObjQty - 1
                n2 = n2 + 1

                If n1 < tObjIndex Then
                    tmpList(n2) = ObjList(n1)
                ElseIf n1 > tObjIndex Then
                    tmpList(n2 - 1) = ObjList(n1)
                End If

            Next n1

            Add2Selection -1
            ObjIndex = 0
            Add2Selection ObjIndex

        Case BringFoward

            If tObjIndex = ObjQty - 1 Then Exit Sub
            tmpList(tObjIndex + 1) = ObjList(tObjIndex)
            tmpList(tObjIndex) = ObjList(tObjIndex + 1)

            For n1 = 0 To ObjQty - 1
                n2 = n2 + 1

                If n1 < tObjIndex Then
                    tmpList(n2 - 1) = ObjList(n1)
                ElseIf n1 > tObjIndex Then
                    tmpList(n2) = ObjList(n1 + 1)
                End If

            Next n1

            Add2Selection -1
            ObjIndex = tObjIndex + 1
            Add2Selection ObjIndex

        Case SendBackward

            If tObjIndex = 0 Then Exit Sub
            tmpList(tObjIndex - 1) = ObjList(tObjIndex)
            tmpList(tObjIndex) = ObjList(tObjIndex - 1)

            For n1 = 0 To ObjQty - 1
                n2 = n2 + 1

                If n1 < tObjIndex - 1 Then
                    tmpList(n2 - 1) = ObjList(n1)
                ElseIf n1 >= tObjIndex Then
                    tmpList(n2) = ObjList(n1 + 1)
                End If

            Next n1

            Add2Selection -1
            ObjIndex = tObjIndex - 1
            Add2Selection ObjIndex
    End Select

    For n1 = 0 To ObjQty - 1
        ObjList(n1) = tmpList(n1)
    Next n1

    ReDraw
    Add2UndoBuffer
End Sub

Private Sub ShowSelection()
    SuperDebug "sub/fun: ShowSelection"
    Dim n   As Long
    Dim tC  As myCoorType
    Dim tGr As Integer
    ' show selection dot

    ReDraw False

    For n = 0 To 3
        griff(n).Visible = False
    Next n

    If ObjIndex = -1 Then Exit Sub

    tGr = ObjList(ObjIndex).mGroupMember

    DrawControl.DrawMode = 7

    If tGr = 0 Then

        For n = 0 To QtySel - 1
            tC = GetSelPosition(ObjList(ListSel(n)).mLeft * mZF, ObjList(ListSel(n)).mTop * mZF, ObjList(ListSel(n)).mWidth * mZF, ObjList(ListSel(n)).mHeight * mZF, ObjList(ListSel(n)).mAngle)

            If ListSel(n) = ObjIndex Then
    
                DrawControl.Line (tC.posX1 - 10, tC.posY1 - 10)-(tC.posX1 - 2, tC.posY1 - 2), vbWhite, BF
                DrawControl.Line (tC.posX1 + tC.centerX - 4, tC.posY1 - 10)-(tC.posX1 + tC.centerX + 4, tC.posY1 - 2), vbWhite, BF
                DrawControl.Line (tC.posX4 + 2, tC.posY4 - 10)-(tC.posX4 + 10, tC.posY4 - 2), vbWhite, BF
    
                DrawControl.Line (tC.posX1 - 10, tC.posY1 + tC.centerY - 4)-(tC.posX1 - 2, tC.posY1 + tC.centerY + 4), vbWhite, BF
                DrawControl.Line (tC.posX4 + 2, tC.posY1 + tC.centerY - 4)-(tC.posX4 + 10, tC.posY1 + tC.centerY + 4), vbWhite, BF
    
                DrawControl.Line (tC.posX2 - 10, tC.posY2 + 2)-(tC.posX2 - 2, tC.posY2 + 10), vbWhite, BF
                DrawControl.Line (tC.posX1 + tC.centerX - 4, tC.posY2 + 2)-(tC.posX1 + tC.centerX + 4, tC.posY2 + 10), vbWhite, BF
                DrawControl.Line (tC.posX3 + 2, tC.posY3 + 2)-(tC.posX3 + 10, tC.posY3 + 10), vbWhite, BF
            Else
                DrawControl.Line (tC.posX1 - 10, tC.posY1 - 10)-(tC.posX1 - 2, tC.posY1 - 2), vbWhite, B
                DrawControl.Line (tC.posX1 + tC.centerX - 4, tC.posY1 - 10)-(tC.posX1 + tC.centerX + 4, tC.posY1 - 2), vbWhite, B
                DrawControl.Line (tC.posX4 + 2, tC.posY4 - 10)-(tC.posX4 + 10, tC.posY4 - 2), vbWhite, B
    
                DrawControl.Line (tC.posX1 - 10, tC.posY1 + tC.centerY - 4)-(tC.posX1 - 2, tC.posY1 + tC.centerY + 4), vbWhite, B
                DrawControl.Line (tC.posX4 + 2, tC.posY1 + tC.centerY - 4)-(tC.posX4 + 10, tC.posY1 + tC.centerY + 4), vbWhite, B
    
                DrawControl.Line (tC.posX2 - 10, tC.posY2 + 2)-(tC.posX2 - 2, tC.posY2 + 10), vbWhite, B
                DrawControl.Line (tC.posX1 + tC.centerX - 4, tC.posY2 + 2)-(tC.posX1 + tC.centerX + 4, tC.posY2 + 10), vbWhite, B
                DrawControl.Line (tC.posX3 + 2, tC.posY3 + 2)-(tC.posX3 + 10, tC.posY3 + 10), vbWhite, B
            End If

        Next n

    ElseIf tGr > 0 Then

        For n = 0 To ObjQty - 1

            If ObjList(n).mGroupMember = tGr Then
                tC = GetSelPosition(ObjList(n).mLeft * mZF, ObjList(n).mTop * mZF, ObjList(n).mWidth * mZF, ObjList(n).mHeight * mZF, ObjList(n).mAngle)
    
                DrawControl.Line (tC.posX1 - 10, tC.posY1 - 10)-(tC.posX1 - 2, tC.posY1 - 2), vbYellow, B
                DrawControl.Line (tC.posX1 + tC.centerX - 4, tC.posY1 - 10)-(tC.posX1 + tC.centerX + 4, tC.posY1 - 2), vbYellow, B
                DrawControl.Line (tC.posX4 + 2, tC.posY4 - 10)-(tC.posX4 + 10, tC.posY4 - 2), vbYellow, B
    
                DrawControl.Line (tC.posX1 - 10, tC.posY1 + tC.centerY - 4)-(tC.posX1 - 2, tC.posY1 + tC.centerY + 4), vbYellow, B
                DrawControl.Line (tC.posX4 + 2, tC.posY1 + tC.centerY - 4)-(tC.posX4 + 10, tC.posY1 + tC.centerY + 4), vbYellow, B
    
                DrawControl.Line (tC.posX2 - 10, tC.posY2 + 2)-(tC.posX2 - 2, tC.posY2 + 10), vbYellow, B
                DrawControl.Line (tC.posX1 + tC.centerX - 4, tC.posY2 + 2)-(tC.posX1 + tC.centerX + 4, tC.posY2 + 10), vbYellow, B
                DrawControl.Line (tC.posX3 + 2, tC.posY3 + 2)-(tC.posX3 + 10, tC.posY3 + 10), vbYellow, B
            End If

        Next n

    End If

    DrawControl.DrawMode = 13

    With ObjList(ObjIndex)
        RaiseEvent ObjSelected(.mObjectType, .mGISObj, ObjIndex, .mLeft, .mTop, .mWidth, .mHeight, .mAngle, .mFillColor, .mFillStyle, .mBorderColor, .mBorderWidth, .mAspect, .mFontName, .mFontSize, .mFontBold, .mFontItalic, .mFontUnderline, .mFontStrikethru, .mText, .mTextAlign, .mPointQty)
    End With

End Sub

Public Sub AlignSelectedObjects(cAlign As myAlignType)
    SuperDebug "sub/fun: AlignSelectedObjects"
    Dim n    As Long
    Dim tmp  As Single
    Dim minX As Single
    Dim maxX As Single
    Dim minY As Single
    Dim maxY As Single
    Dim tC   As myCoorType

    NextLine = False

    NewObj = False

    If QtySel = 0 Or ObjIndex = -1 Then Exit Sub

    tC = GetCoordinate(ObjList(ObjIndex).mLeft, ObjList(ObjIndex).mTop, ObjList(ObjIndex).mWidth, ObjList(ObjIndex).mHeight, ObjList(ObjIndex).mAngle)

    Select Case ObjList(ObjIndex).mAngle

        Case 0 To 90
            minX = tC.posX1
            maxX = tC.posX4
            minY = tC.posY3
            maxY = tC.posY2

        Case 91 To 180
            minX = tC.posX3
            maxX = tC.posX2
            minY = tC.posY4
            maxY = tC.posY1

        Case 181 To 270
            minX = tC.posX4
            maxX = tC.posX1
            minY = tC.posY4
            maxY = tC.posY1

        Case 271 To 360
            minX = tC.posX2
            maxX = tC.posX3
            minY = tC.posY1
            maxY = tC.posY4
    End Select

    For n = 0 To QtySel - 1

        Select Case cAlign

            Case mLeft
                ObjList(ListSel(n)).mLeft = minX

            Case mCenter_V
                tmp = ((maxX - minX) / 2) + minX
                ObjList(ListSel(n)).mLeft = tmp - (ObjList(ListSel(n)).mWidth / 2)

            Case mRight
                ObjList(ListSel(n)).mLeft = maxX - ObjList(ListSel(n)).mWidth

            Case mTop
                ObjList(ListSel(n)).mTop = minY

            Case mCenter_H
                tmp = ((maxY - minY) / 2) + minY
                ObjList(ListSel(n)).mTop = tmp - (ObjList(ListSel(n)).mHeight / 2)

            Case mBottom
                ObjList(ListSel(n)).mTop = maxY - ObjList(ListSel(n)).mHeight

            Case mCenter_V_H
                tmp = ((maxX - minX) / 2) + minX
                ObjList(ListSel(n)).mLeft = tmp - (ObjList(ListSel(n)).mWidth / 2)
                tmp = ((maxY - minY) / 2) + minY
                ObjList(ListSel(n)).mTop = tmp - (ObjList(ListSel(n)).mHeight / 2)

            Case mCenterPage
                tmp = mCanvasWidth / 2
                ObjList(ListSel(n)).mLeft = tmp - (ObjList(ListSel(n)).mWidth / 2)
                tmp = mCanvasHeight / 2
                ObjList(ListSel(n)).mTop = tmp - (ObjList(ListSel(n)).mHeight / 2)
        End Select

    Next n

    ReDraw
    Add2UndoBuffer
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    SuperDebug "sub/fun: UserControl_KeyDown"
    If Shift = 0 Then

        Select Case KeyCode

            Case vbKeyAdd
                mZF = mZF + 0.1

                If mZF > 10 Then mZF = 10
                UserControl_Resize
                ReDraw

            Case vbKeySubtract
                mZF = mZF - 0.1

                If mZF < 0.1 Then mZF = 0.1
                UserControl_Resize
                ReDraw
        End Select

    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SuperDebug "sub/fun: UserControl_ReadProperties"
    With PropBag
        DrawControl.BackColor = .ReadProperty("BackColor", vbWhite)
        mDefaultText = .ReadProperty("DefaultText", "New Text")
        mCanvasWidth = .ReadProperty("CanvasWidth", 640)
        mCanvasHeight = .ReadProperty("CanvasHeight", 480)
        mUndoSize = .ReadProperty("UndoBufferSize", 10)
        mShowCanvasSize = .ReadProperty("ShowCanvasSize", True)
        mArrowStep = .ReadProperty("ArrowStep", 1)
        mZF = .ReadProperty("ZoomFactor", 1)
    End With

    ReDim UndoBuffer(mUndoSize)
End Sub

Private Sub UserControl_Resize()
    SuperDebug "sub/fun: UserControl_Resize"
    Dim uW As Long
    Dim uH As Long

    DrawControl.Width = mCanvasWidth * mZF
    DrawControl.Height = mCanvasHeight * mZF

    uW = UserControl.ScaleX(UserControl.Width, vbTwips, vbPixels) - 4
    uH = UserControl.ScaleY(UserControl.Height, vbTwips, vbPixels) - 4

    DrawControl.left = (uW / 2 - (mCanvasWidth * mZF / 2))
    DrawControl.top = (uH / 2 - (mCanvasHeight * mZF / 2))

    HScroll1.Visible = False
    VScroll1.Visible = False
    Corner.Visible = False

    If DrawControl.Width + 18 > uW Then
        HScroll1.left = 0
        HScroll1.top = uH - 16
        HScroll1.Width = uW
        HScroll1.Visible = True
        HScroll1.Max = -(DrawControl.Width - uW) - 40
        HScroll1.Min = 20

        If toZoom = False Then
            HScroll1_Change
            HScroll1.value = 20
        End If
    End If

    If DrawControl.Height + 18 > uH Then
        VScroll1.left = uW - 16
        VScroll1.top = 0
        VScroll1.Height = uH
        VScroll1.Visible = True
        VScroll1.Max = -(DrawControl.Height - uH) - 40
        VScroll1.Min = 20

        If toZoom = False Then
            VScroll1_Change
            VScroll1.value = 20
        End If
    End If

    toZoom = False

    If DrawControl.Width + 18 > uW And DrawControl.Height + 18 > uH Then
        HScroll1.Width = uW - 16
        VScroll1.Height = uH - 16
        Corner.left = uW - 16
        Corner.top = uH - 16
        Corner.Visible = True
    Else
        UserControl.Cls
        UserControl.FontBold = True
        UserControl.DrawWidth = 1
        UserControl.Line (DrawControl.left + 4, DrawControl.top + 4)-Step(DrawControl.Width + 2, DrawControl.Height + 2), &H80000015, BF
        UserControl.Line (DrawControl.left - 1, DrawControl.top - 1)-Step(DrawControl.Width + 1, DrawControl.Height + 1), , B

        If mShowCanvasSize = True Then
            UserControl.CurrentX = DrawControl.left + DrawControl.Width - UserControl.TextWidth(mCanvasWidth & " X " & mCanvasHeight) + 7
            UserControl.CurrentY = DrawControl.top + DrawControl.Height + 7
            UserControl.Print mCanvasWidth & " X " & mCanvasHeight
        End If
    End If

End Sub

Private Sub UserControl_Show()
    SuperDebug "sub/fun: UserControl_Show"
    RaiseEvent UndoRedo(True, True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    SuperDebug "sub/fun: UserControl_WriteProperties"
    With PropBag
        .WriteProperty "BackColor", DrawControl.BackColor, vbWhite
        .WriteProperty "DefaultText", mDefaultText, "New Text"
        .WriteProperty "CanvasWidth", mCanvasWidth, 640
        .WriteProperty "CanvasHeight", mCanvasHeight, 480
        .WriteProperty "UndoBufferSize", mUndoSize, 10
        .WriteProperty "ShowCanvasSize", mShowCanvasSize, True
        .WriteProperty "ArrowStep", mArrowStep, 1
        .WriteProperty "ZoomFactor", mZF, 1
    End With

End Sub

Public Property Get DefaultText() As String
    DefaultText = mDefaultText
End Property

Public Property Let DefaultText(ByVal sDefaultText As String)
    SuperDebug "sub/fun: DefaultText"
    mDefaultText = sDefaultText
    PropertyChanged "DefaultText"
End Property

Public Sub NewProject()
    SuperDebug "sub/fun: NewProject"
    ReDim ObjList(0)
    ObjQty = 0
    ReDim ListSel(0)
    QtySel = 0
    ObjIndex = -1
    GroupQty = 0

    NewObj = False
    isDown = False
    isMove = False
    onObject = False
    isResize = False
    toSize = -1
    NewText = False

    If isUndo = False Then
        RaiseEvent UndoRedo(True, True)
        ReDim UndoBuffer(mUndoSize)
        UndoStack = 0
        UndoPointer = 0
    End If

    ReDraw
End Sub

Public Sub UseSelector()
    SuperDebug "sub/fun: UseSelector"
    Add2Selection -1

    DrawControl.MousePointer = 0

    If NewObj = True Or NewText = True Then
        ObjQty = ObjQty - 1
        ReDim Preserve ObjList(ObjQty)
    End If

    NewObj = False
    isDown = False
    isMove = False
    onObject = False
    isResize = False
    toSize = -1
    NewText = False
    myText.Visible = False
    'ObjIndex = -1
    ReDraw
End Sub

Public Property Get CanvasWidth() As Long
    CanvasWidth = mCanvasWidth
End Property

Public Property Let CanvasWidth(ByVal lCanvasWidth As Long)
    SuperDebug "sub/fun: CanvasWidth"
    mCanvasWidth = lCanvasWidth
    UserControl_Resize
    PropertyChanged "CanvasWidth"
End Property

Public Property Get CanvasHeight() As Long
    CanvasHeight = mCanvasHeight
End Property

Public Property Let CanvasHeight(ByVal lCanvasHeight As Long)
    SuperDebug "sub/fun: CanvasHeight"
    mCanvasHeight = lCanvasHeight
    UserControl_Resize
    PropertyChanged "CanvasHeight"
End Property

Public Property Get CanvasCenterX() As Long
    CanvasCenterX = mCanvasWidth / 2
End Property

Public Property Get CanvasCenterY() As Long
    CanvasCenterY = mCanvasHeight / 2
End Property

Private Sub VScroll1_Change()
    SuperDebug "sub/fun: VScroll1_Change"
    On Error Resume Next
    DrawControl.top = VScroll1.value

    UserControl.Cls
    UserControl.DrawWidth = 1
    UserControl.Line (DrawControl.left + 4, DrawControl.top + 4)-Step(DrawControl.Width + 2, DrawControl.Height + 2), &H80000015, BF
    UserControl.Line (DrawControl.left - 1, DrawControl.top - 1)-Step(DrawControl.Width + 1, DrawControl.Height + 1), , B

    If mShowCanvasSize = True Then
        UserControl.CurrentX = DrawControl.left + DrawControl.Width - UserControl.TextWidth(mCanvasWidth & " X " & mCanvasHeight) + 7
        UserControl.CurrentY = DrawControl.top + DrawControl.Height + 7
        UserControl.Print mCanvasWidth & " X " & mCanvasHeight
    End If

    DrawControl.SetFocus
End Sub

Public Sub CopyObject()
    SuperDebug "sub/fun: CopyObject"
    Dim n   As Long
    Dim tGr As Integer

    If QtySel > 0 Then
        tGr = ObjList(ObjIndex).mGroupMember

        If tGr = 0 Then
            ClipQty = QtySel
            ReDim mClipBoard(QtySel)

            For n = 0 To QtySel - 1
                mClipBoard(n) = ObjList(ListSel(n))
            Next n

        ElseIf tGr > 0 Then
            ClipQty = 0

            For n = 0 To ObjQty - 1

                If ObjList(n).mGroupMember = tGr Then
                    ClipQty = ClipQty + 1
                    ReDim Preserve mClipBoard(ClipQty)
                    mClipBoard(ClipQty - 1) = ObjList(n)
                End If

            Next n

        End If

    Else
        ClipQty = 0
    End If

End Sub

Public Sub PasteObject()
    SuperDebug "sub/fun: PasteObject"
    Dim n   As Long
    Dim tGr As Integer

    tGr = mClipBoard(0).mGroupMember

    If tGr > 0 Then
        GroupQty = GroupQty + 1
        tGr = GroupQty
    End If

    For n = 0 To ClipQty - 1

        With mClipBoard(n)
            AddObject .mObjectType, .mTop, .mLeft, .mHeight, .mWidth, .mAngle, .mFillColor, .mFillStyle, .mBorderColor, .mBorderWidth, .mPicture, .mFontName, .mFontSize, .mFontBold, .mFontItalic, .mFontUnderline, .mFontStrikethru, .mText, .mTextAlign, .mPointQty, .mPosX0, .mPosY0, .mPosX1, .mPosY1, .mPosX2, .mPosY2, .mPosX3, .mPosY3, tGr, .mAspect
        End With

    Next n

    Add2UndoBuffer
End Sub

Public Sub Undo()

    SuperDebug "sub/fun: Undo"
    If UndoPointer < mUndoSize - 1 Then

        UndoPointer = UndoPointer + 1
        isUndo = True
        RestoreUndo UndoPointer
        isUndo = False

        If UndoPointer = UndoStack Then
            RaiseEvent UndoRedo(True, False)
        Else
            RaiseEvent UndoRedo(False, Not CBool(UndoStack))
        End If

    Else
        RaiseEvent UndoRedo(True, Not CBool(UndoStack))
    End If

End Sub

Public Sub Redo()

    SuperDebug "sub/fun: Redo"
    If UndoPointer > 1 Then
        isUndo = True
        UndoPointer = UndoPointer - 1
        RestoreUndo UndoPointer

        isUndo = False
        RaiseEvent UndoRedo(False, False)
    ElseIf UndoPointer = 1 Then
        isUndo = True
        UndoPointer = UndoPointer - 1
        RestoreUndo UndoPointer
        isUndo = False
        RaiseEvent UndoRedo(False, True)
    Else
        RaiseEvent UndoRedo(False, True)
    End If

End Sub

Public Property Get UndoBufferSize() As Integer
    UndoBufferSize = mUndoSize
End Property

Public Property Let UndoBufferSize(ByVal iUndoBufferSize As Integer)
    SuperDebug "sub/fun: UndoBufferSize"
    mUndoSize = iUndoBufferSize
    ReDim UndoBuffer(mUndoSize)
    Call UserControl.PropertyChanged("UndoBufferSize")
End Property

Private Sub Add2UndoBuffer()
    SuperDebug "sub/fun: Add2UndoBuffer"
    Dim n       As Long
    Dim FF      As Integer
    Dim mData   As String
    Dim tmpFile As String

    If isUndo = True Then Exit Sub

    For n = 0 To ObjQty - 1

        With ObjList(n)
            mData = mData & String(5, 254) & String(5, 255) & String(5, 254) & .mObjectType & Chr(0) & Hex(.mTop) & Chr(0) & Hex(.mLeft) & Chr(0) & Hex(.mHeight) & Chr(0) & Hex(.mWidth) & Chr(0) & Hex(.mAngle) & Chr(0) & Hex(DrawControl.BackColor) & Chr(0) & Hex(.mFillColor) & Chr(0) & Hex(.mFillStyle) & Chr(0) & Hex(.mBorderColor) & Chr(0) & Hex(.mBorderWidth) & Chr(0) & Hex(.mAspect) & Chr(0) & .mFontName & Chr(0) & .mFontSize & Chr(0) & Int(.mFontBold) & Chr(0) & Int(.mFontItalic) & Chr(0) & Int(.mFontUnderline) & Chr(0) & Int(.mFontStrikethru) & Chr(0) & .mText & Chr(0) & .mTextAlign & Chr(0) & Hex(.mPointQty) & Chr(0) & Hex(.mPosX0) & Chr(0) & Hex(.mPosY0) & Chr(0) & Hex(.mPosX1) & Chr(0) & Hex(.mPosY1) & Chr(0) & Hex(.mPosX2) & Chr(0) & Hex(.mPosY2) & Chr(0) & Hex(.mPosX3) & Chr(0) & Hex(.mPosY3) & Chr(0) & Hex(.mGroupMember) & Chr(0) & Hex(mCanvasWidth) & Chr(0) & Hex(mCanvasHeight) & Chr(0) & Int(.mDeleted)

            If ObjList(n).mObjectType = mImage Then
                Set PicData.Picture = ObjList(n).mPicture
                SavePicture PicData, App.Path & "\Tmp.bmp"

                DoEvents
                FF = FreeFile
                Open App.Path & "\Tmp.bmp" For Binary As #FF
                tmpFile = Input(FileLen(App.Path & "\Tmp.bmp"), #FF)
                Close #FF

                DoEvents
                mData = mData & String(5, 255) & String(5, 254) & String(5, 255) & tmpFile

                If Len(Dir(App.Path & "\Tmp.bmp")) Then Kill App.Path & "\Tmp.bmp"

                DoEvents
            End If

        End With

    Next n

    For n = mUndoSize - 1 To 1 Step -1
        UndoBuffer(n) = UndoBuffer(n - 1)
    Next n

    UndoBuffer(0) = mData
    UndoPointer = 0

    If UndoStack < mUndoSize Then UndoStack = UndoStack + 1
    RaiseEvent UndoRedo(False, Not CBool(UndoPointer))
    RaiseEvent Prompt2Save

End Sub

Private Sub RestoreUndo(Pointer As Integer)
    SuperDebug "sub/fun: RestoreUndo"
    Dim n      As Long
    Dim FF     As Integer
    Dim mFile  As String
    Dim mData  As String
    Dim tmpPic As String

    If Len(UndoBuffer(Pointer)) > 0 Then
        ObjQty = UBound(Split(UndoBuffer(Pointer), String(5, 254) & String(5, 255) & String(5, 254)))

        ReDim ObjList(ObjQty)

        For n = 0 To ObjQty - 1
            mFile = Split(UndoBuffer(Pointer), String(5, 254) & String(5, 255) & String(5, 254))(n + 1)
            mData = Split(mFile, String(5, 255) & String(5, 254) & String(5, 255))(0)

            With ObjList(n)
                .mObjectType = Int(Split(mData, Chr(0))(0))
                .mTop = CLng("&H" & Split(mData, Chr(0))(1))
                .mLeft = CLng("&H" & Split(mData, Chr(0))(2))
                .mHeight = CLng("&H" & Split(mData, Chr(0))(3))
                .mWidth = CLng("&H" & Split(mData, Chr(0))(4))
                .mAngle = CLng("&H" & Split(mData, Chr(0))(5))
                DrawControl.BackColor = CLng("&H" & Split(mData, Chr(0))(6))
                .mFillColor = CLng("&H" & Split(mData, Chr(0))(7))
                .mFillStyle = CLng("&H" & Split(mData, Chr(0))(8))
                .mBorderColor = CLng("&H" & Split(mData, Chr(0))(9))
                .mBorderWidth = CLng("&H" & Split(mData, Chr(0))(10))
                .mAspect = CLng("&H" & Split(mData, Chr(0))(11))
                .mFontName = Split(mData, Chr(0))(12)
                .mFontSize = Split(mData, Chr(0))(13)
                .mFontBold = CBool(Split(mData, Chr(0))(14))
                .mFontItalic = CBool(Split(mData, Chr(0))(15))
                .mFontUnderline = CBool(Split(mData, Chr(0))(16))
                .mFontStrikethru = CBool(Split(mData, Chr(0))(17))
                .mText = Split(mData, Chr(0))(18)
                .mTextAlign = Split(mData, Chr(0))(19)
                .mPointQty = CLng("&H" & Split(mData, Chr(0))(20))
                .mPosX0 = CLng("&H" & Split(mData, Chr(0))(21))
                .mPosY0 = CLng("&H" & Split(mData, Chr(0))(22))
                .mPosX1 = CLng("&H" & Split(mData, Chr(0))(23))
                .mPosY1 = CLng("&H" & Split(mData, Chr(0))(24))
                .mPosX2 = CLng("&H" & Split(mData, Chr(0))(25))
                .mPosY2 = CLng("&H" & Split(mData, Chr(0))(26))
                .mPosX3 = CLng("&H" & Split(mData, Chr(0))(27))
                .mPosY3 = CLng("&H" & Split(mData, Chr(0))(28))
                .mGroupMember = CLng("&H" & Split(mData, Chr(0))(29))
                mCanvasWidth = CLng("&H" & Split(mData, Chr(0))(30))
                mCanvasHeight = CLng("&H" & Split(mData, Chr(0))(31))
                .mDeleted = CBool(Split(mData, Chr(0))(32))

                If .mObjectType = mImage Then
                    tmpPic = Split(mFile, String(5, 255) & String(5, 254) & String(5, 255))(1)
                    FF = FreeFile
                    Open App.Path & "\Tmp.bmp" For Binary Access Write As #FF
                    Put #FF, , tmpPic
                    Close #FF

                    DoEvents
                    Set .mPicture = LoadPicture(App.Path & "\Tmp.bmp")

                    DoEvents

                    If Len(Dir(App.Path & "\Tmp.bmp")) Then Kill App.Path & "\Tmp.bmp"

                    DoEvents
                End If

            End With

        Next n

        mZF = 1
        RedimList
        UserControl_Resize
        ReDraw
    Else
        NewProject
    End If

End Sub

Public Sub SelectAllObjects()
    SuperDebug "sub/fun: SelectAllObjects"
    Dim n As Long

    QtySel = ObjQty

    ReDim ListSel(QtySel)

    For n = 0 To QtySel - 1
        ListSel(n) = n
    Next n

    ObjIndex = QtySel - 1
    ReDraw
End Sub

Public Property Get ShowCanvasSize() As Boolean
    ShowCanvasSize = mShowCanvasSize
End Property

Public Property Let ShowCanvasSize(ByVal bShowCanvasSize As Boolean)
    SuperDebug "sub/fun: ShowCanvasSize"
    mShowCanvasSize = bShowCanvasSize
    UserControl_Resize
    PropertyChanged "ShowCanvasSize"
End Property

Private Sub DrawRectangle(cLeft As Single, _
                          cTop As Single, _
                          cWidth As Single, _
                          cHeight As Single, _
                          Optional cAngle As Single)
    SuperDebug "sub/fun: DrawRectangle"
    Dim tCoor         As myCoorType
    Dim POINT(1 To 4) As POINTAPI

    tCoor = GetCoordinate(cLeft, cTop, cWidth, cHeight, cAngle)

    With tCoor
        POINT(1).X = .posX1
        POINT(1).Y = .posY1
        POINT(2).X = .posX2
        POINT(2).Y = .posY2
        POINT(3).X = .posX4
        POINT(3).Y = .posY4
        POINT(4).X = .posX3
        POINT(4).Y = .posY3
    End With

    Polygon DrawControl.hdc, POINT(1), 4

End Sub

Private Function GetCoordinate(cLeft As Single, _
                               cTop As Single, _
                               cWidth As Single, _
                               cHeight As Single, _
                               cAngle As Single) As myCoorType
    SuperDebug "sub/fun: GetCoordinate"
    Dim rXg1 As Double
    Dim rYg1 As Double
    Dim rXg2 As Double
    Dim rYg2 As Double

    rXg1 = cLeft + (cHeight * Sin((cAngle Mod 360) * Pi / 180))
    rYg1 = cTop + (cHeight * Cos((cAngle Mod 360) * Pi / 180))
    rXg2 = cLeft - (cWidth * Sin((cAngle - 90 Mod 360) * Pi / 180))
    rYg2 = cTop - (cWidth * Cos((cAngle - 90 Mod 360) * Pi / 180))

    With GetCoordinate
        .posX1 = cLeft
        .posY1 = cTop
        .posX2 = rXg1
        .posY2 = rYg1
        .posX3 = rXg2
        .posY3 = rYg2
        .posX4 = rXg2 + (rXg1 - cLeft)
        .posY4 = rYg2 + (rYg1 - cTop)
        .centerX = Abs(((rXg2 - rXg1) / 2) + rXg1)
        .centerY = Abs(((rYg2 - rYg1) / 2) + rYg1)
    End With

End Function

Private Function GetSelPosition(cLeft As Single, _
                                cTop As Single, _
                                cWidth As Single, _
                                cHeight As Single, _
                                cAngle As Single) As myCoorType
    SuperDebug "sub/fun: GetSelPosition"
    Dim rx(1 To 4) As Double
    Dim ry(1 To 4) As Double
    Dim n          As Integer
    Dim xmin       As Single
    Dim xmax       As Single
    Dim ymin       As Single
    Dim ymax       As Single

    rx(1) = cLeft
    ry(1) = cTop
    rx(2) = cLeft + (cHeight * Sin((cAngle Mod 360) * Pi / 180))
    ry(2) = cTop + (cHeight * Cos((cAngle Mod 360) * Pi / 180))
    rx(3) = cLeft - (cWidth * Sin((cAngle - 90 Mod 360) * Pi / 180))
    ry(3) = cTop - (cWidth * Cos((cAngle - 90 Mod 360) * Pi / 180))
    rx(4) = rx(3) + (rx(2) - rx(1))
    ry(4) = ry(3) + (ry(2) - ry(1))

    xmin = mCanvasWidth
    xmax = 0
    ymin = mCanvasHeight
    ymax = 0

    For n = 1 To 4

        If rx(n) < xmin Then xmin = rx(n)
        If rx(n) > xmax Then xmax = rx(n)
        If ry(n) < ymin Then ymin = ry(n)
        If ry(n) > ymax Then ymax = ry(n)
    Next n

    With GetSelPosition
        .posX1 = xmin
        .posY1 = ymin
        .posX2 = xmin
        .posY2 = ymax
        .posX3 = xmax
        .posY3 = ymax
        .posX4 = xmax
        .posY4 = ymin
        .centerX = (xmax - xmin) / 2
        .centerY = (ymax - ymin) / 2
    End With

End Function

Private Sub DrawText(Txt As String, _
                     cLeft As Single, _
                     cTop As Single, _
                     cWidth As Single, _
                     cHeight As Single, _
                     FontName As String, _
                     ByVal cSize As Double, _
                     ByVal Angle As Single, _
                     ByVal Bold As Boolean, _
                     ByVal Italic As Boolean, _
                     ByVal Underline As Boolean, _
                     ByVal Strikethru As Boolean, _
                     Optional cAlign As AlignmentConstants = vbLeftJustify)
    SuperDebug "sub/fun: DrawText"
    Dim newFont     As Long
    Dim Oldfont     As Long
    Dim nEscapement As Long
    Dim nHeight     As Long
    Dim Weight      As FontWeight

    If Bold = True Then Weight = FW_BOLD Else Weight = FW_NORMAL

    nEscapement = Angle * 10

    DrawControl.FontName = FontName
    DrawControl.FontSize = cSize
    DrawControl.FontBold = Bold
    DrawControl.FontItalic = Italic
    DrawControl.FontUnderline = Underline
    DrawControl.FontStrikethru = Strikethru

    nHeight = -MulDiv(cSize, GetDeviceCaps(DrawControl.hdc, LOGPIXELSY), 72)

    newFont = CreateFont(nHeight, 0, nEscapement, nEscapement, Weight, Italic, Underline, Strikethru, ANSI_CHARSET, OUT_TT_PRECIS, CLIP_LH_ANGLES, PROOF_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, FontName)

    Oldfont = SelectObject(DrawControl.hdc, newFont)

    PrintWordWrap Txt, cLeft, cTop, cWidth, cHeight, cAlign

    newFont = SelectObject(DrawControl.hdc, Oldfont)

    DeleteObject newFont
End Sub

Public Property Get ArrowStep() As Integer
    ArrowStep = mArrowStep
End Property

Public Property Let ArrowStep(ByVal vNewArrowStep As Integer)
    SuperDebug "sub/fun: ArrowStep"
    mArrowStep = vNewArrowStep
    PropertyChanged "ArrowStep"
End Property

Public Property Get ZoomFactor() As Single
    ZoomFactor = mZF
End Property

Public Property Let ZoomFactor(ByVal sZoomFactor As Single)
    SuperDebug "sub/fun: ZoomFactor"
    mZF = sZoomFactor
    toZoom = True
    UserControl_Resize
    ReDraw
    PropertyChanged "ZoomFactor"
End Property

Public Sub UnSelectAll()
    SuperDebug "sub/fun: UnSelectAll"
    QtySel = 0
    ReDim ListSel(0)
    ObjIndex = -1
    ReDraw
End Sub

Private Sub DrawEllipse(cLeft As Single, _
                        cTop As Single, _
                        cWidth As Single, _
                        cHeight As Single, _
                        Optional cAngle As Single)
    SuperDebug "sub/fun: DrawEllipse"
    Dim Pts() As POINTAPI

    Pts = EllipsePts(cLeft, cTop, cWidth, cHeight, cAngle)

    BeginPath DrawControl.hdc
    PolyBezier DrawControl.hdc, Pts(0), UBound(Pts) + 1
    EndPath DrawControl.hdc
    StrokeAndFillPath DrawControl.hdc
End Sub

Private Sub DrawArc(cIndex As Long, _
                    cLeft As Single, _
                    cTop As Single, _
                    cWidth As Single, _
                    cHeight As Single, _
                    Optional posX0 As Single = -1, _
                    Optional posY0 As Single = -1, _
                    Optional posX1 As Single = -1, _
                    Optional posY1 As Single = -1, _
                    Optional posX2 As Single = -1, _
                    Optional posY2 As Single = -1, _
                    Optional posX3 As Single = -1, _
                    Optional posY3 As Single = -1)

    SuperDebug "sub/fun: DrawArc"
    Dim Pts(1 To 4) As POINTAPI
    Dim n           As Integer
    Dim xmin        As Single
    Dim xmax        As Single
    Dim ymin        As Single
    Dim ymax        As Single

    If posX0 > 0 Then Pts(1).X = posX0 Else Pts(1).X = cLeft
    If posY0 > 0 Then Pts(1).Y = posY0 Else Pts(1).Y = cTop + cHeight
    If posX1 > 0 Then Pts(2).X = posX1 Else Pts(2).X = cLeft
    If posY1 > 0 Then Pts(2).Y = posY1 Else Pts(2).Y = cTop
    If posX2 > 0 Then Pts(3).X = posX2 Else Pts(3).X = cLeft + cWidth
    If posY2 > 0 Then Pts(3).Y = posY2 Else Pts(3).Y = cTop
    If posX3 > 0 Then Pts(4).X = posX3 Else Pts(4).X = cLeft + cWidth
    If posY3 > 0 Then Pts(4).Y = posY3 Else Pts(4).Y = cTop + cHeight

    PolyBezier DrawControl.hdc, Pts(1), 4

    xmin = mCanvasWidth
    xmax = 0
    ymin = mCanvasHeight
    ymax = 0

    For n = 1 To 4

        If Pts(n).X < xmin Then xmin = Pts(n).X
        If Pts(n).X > xmax Then xmax = Pts(n).X
        If Pts(n).Y < ymin Then ymin = Pts(n).Y
        If Pts(n).Y > ymax Then ymax = Pts(n).Y
    Next n

    With ObjList(cIndex)
        .mTop = ymin / mZF
        .mLeft = xmin / mZF
        .mWidth = (xmax - xmin) / mZF
        .mHeight = (ymax - ymin) / mZF
    End With

End Sub

Private Sub Add2Selection(ObjectIndex As Long)
    SuperDebug "sub/fun: Add2Selection"
    Dim n As Long

    If ObjectIndex > -1 Then

        For n = 0 To QtySel - 1

            If ListSel(n) = ObjectIndex Then
                Exit Sub
            End If

        Next n

        QtySel = QtySel + 1
        ReDim Preserve ListSel(QtySel)
        ListSel(QtySel - 1) = ObjectIndex
    Else
        ReDim ListSel(0)
        QtySel = 0
    End If

End Sub

Private Sub DrawPicture(cPicture As StdPicture, _
                        cLeft As Single, _
                        cTop As Single, _
                        cWidth As Single, _
                        cHeight As Single, _
                        Optional cAngle As Single)
    SuperDebug "sub/fun: DrawPicture"
    Dim Points(3)    As POINTAPI
    Dim DefPoints(3) As POINTAPI
    Dim TS           As Single
    Dim tC           As Single
    Dim SrcWidth     As Single
    Dim SrcHeight    As Single
    Dim wScale       As Single
    Dim hScale       As Single

    SrcWidth = DrawControl.ScaleX(cPicture.Width)
    SrcHeight = DrawControl.ScaleY(cPicture.Height)

    wScale = cWidth / SrcWidth
    hScale = cHeight / SrcHeight

    Points(0).X = 0
    Points(0).Y = 0

    Points(1).X = Points(0).X + SrcWidth * wScale
    Points(1).Y = Points(0).Y

    Points(2).X = Points(0).X
    Points(2).Y = Points(0).Y + SrcHeight * hScale

    TS = Sin(-cAngle * Pi / 180)
    tC = Cos(-cAngle * Pi / 180)

    DefPoints(0).X = (Points(0).X * tC - Points(0).Y * TS) + cLeft
    DefPoints(0).Y = (Points(0).X * TS + Points(0).Y * tC) + cTop

    DefPoints(1).X = (Points(1).X * tC - Points(1).Y * TS) + cLeft
    DefPoints(1).Y = (Points(1).X * TS + Points(1).Y * tC) + cTop

    DefPoints(2).X = (Points(2).X * tC - Points(2).Y * TS) + cLeft
    DefPoints(2).Y = (Points(2).X * TS + Points(2).Y * tC) + cTop

    Set PicData.Picture = cPicture
    PlgBlt DrawControl.hdc, DefPoints(0), PicData.hdc, 0, 0, SrcWidth, SrcHeight, 0, 0, 0
    Set PicData.Picture = Nothing
End Sub

Private Function PolygonPts(cPtsQty As Integer, _
                            cLeft As Single, _
                            cTop As Single, _
                            cWidth As Single, _
                            cHeight As Single, _
                            cAngle As Single) As POINTAPI()
    SuperDebug "sub/fun: PolygonPts"
    Dim POINT()  As POINTAPI
    Dim n        As Integer
    Dim RadiusW  As Single
    Dim RadiusH  As Single
    Dim iCounter As Integer
    Dim R        As Single
    Dim Alfa     As Single

    RadiusW = cWidth / 2
    RadiusH = cHeight / 2

    ReDim POINT(cPtsQty)

    For n = 0 To 360 Step 360 / cPtsQty
        POINT(iCounter).X = RadiusW + Sin(n * Pi / 180) * RadiusW
        POINT(iCounter).Y = RadiusH + Cos(n * Pi / 180) * RadiusH
        R = Sqr(POINT(iCounter).X ^ 2 + POINT(iCounter).Y ^ 2)
        Alfa = Atn2(POINT(iCounter).Y, POINT(iCounter).X) - (cAngle * Pi / 180)
        POINT(iCounter).X = cLeft + R * Cos(Alfa)
        POINT(iCounter).Y = cTop + R * Sin(Alfa)
        iCounter = iCounter + 1
    Next

    PolygonPts = POINT

End Function

Private Function StarPts(cPtsQty As Integer, _
                         cLeft As Single, _
                         cTop As Single, _
                         cWidth As Single, _
                         cHeight As Single, _
                         cAngle As Single) As POINTAPI()
    SuperDebug "sub/fun: StarPts"
    Dim POINT()  As POINTAPI
    Dim n        As Integer
    Dim RadiusW  As Single
    Dim RadiusH  As Single
    Dim iCounter As Integer
    Dim R        As Single
    Dim Alfa     As Single
    Dim isIn     As Boolean

    RadiusW = cWidth / 2
    RadiusH = cHeight / 2

    ReDim POINT(cPtsQty)

    For n = 0 To 360 Step 360 / cPtsQty

        If isIn = False Then
            POINT(iCounter).X = RadiusW + Sin(n * Pi / 180) * RadiusW
            POINT(iCounter).Y = RadiusH + Cos(n * Pi / 180) * RadiusH
        Else
            POINT(iCounter).X = RadiusW + Sin(n * Pi / 180) * RadiusW / 2.5
            POINT(iCounter).Y = RadiusH + Cos(n * Pi / 180) * RadiusH / 2.5
        End If

        isIn = Not isIn
        R = Sqr(POINT(iCounter).X ^ 2 + POINT(iCounter).Y ^ 2)
        Alfa = Atn2(POINT(iCounter).Y, POINT(iCounter).X) - (cAngle * Pi / 180)
        POINT(iCounter).X = cLeft + R * Cos(Alfa)
        POINT(iCounter).Y = cTop + R * Sin(Alfa)
        iCounter = iCounter + 1
    Next

    StarPts = POINT

End Function

Private Sub DrawPolygon(cPtsQty As Integer, _
                        cLeft As Single, _
                        cTop As Single, _
                        cWidth As Single, _
                        cHeight As Single, _
                        Optional cAngle As Single)
    SuperDebug "sub/fun: DrawPolygon"
    Dim Pts() As POINTAPI

    If cPtsQty < 3 Then cPtsQty = 3
    If cPtsQty > 45 Then cPtsQty = 45

    Pts = PolygonPts(cPtsQty, cLeft, cTop, cWidth, cHeight, cAngle)

    Polygon DrawControl.hdc, Pts(0), cPtsQty

End Sub

Private Sub DrawStar(cPtsQty As Integer, _
                     cLeft As Single, _
                     cTop As Single, _
                     cWidth As Single, _
                     cHeight As Single, _
                     Optional cAngle As Single)
    SuperDebug "sub/fun: DrawStar"
    Dim Pts()   As POINTAPI
    Dim tPtsQty As Integer

    If cPtsQty < 3 Then cPtsQty = 3
    If cPtsQty > 30 Then cPtsQty = 30

    tPtsQty = cPtsQty * 2

    Pts = StarPts(tPtsQty, cLeft, cTop, cWidth, cHeight, cAngle)

    Polygon DrawControl.hdc, Pts(0), tPtsQty

End Sub

Public Sub GroupObjects()
    SuperDebug "sub/fun: GroupObjects"
    Dim n As Integer

    If QtySel > 1 Then
        GroupQty = GroupQty + 1

        For n = 0 To QtySel - 1
            ObjList(ListSel(n)).mGroupMember = GroupQty
        Next n

        ReDraw
    End If

End Sub

Public Sub UnGroupObjects()
    SuperDebug "sub/fun: UnGroupObjects"
    Dim n   As Integer
    Dim tGr As Integer

    If ObjIndex > -1 Then
        tGr = ObjList(ObjIndex).mGroupMember

        For n = 0 To ObjQty - 1

            If ObjList(n).mGroupMember = tGr Then ObjList(n).mGroupMember = 0
        Next n

        ReDraw
    End If

End Sub

Private Sub DragBezier(Index As Integer, cx As Single, cy As Single)

    SuperDebug "sub/fun: DragBezier"
    griff(Index).left = cx
    griff(Index).top = cy

    With ObjList(ObjIndex)
        .mPosX0 = griff(0).left / mZF
        .mPosY0 = griff(0).top / mZF
        .mPosX1 = griff(1).left / mZF
        .mPosY1 = griff(1).top / mZF
        .mPosX2 = griff(2).left / mZF
        .mPosY2 = griff(2).top / mZF
        .mPosX3 = griff(3).left / mZF
        .mPosY3 = griff(3).top / mZF
    End With

    ReDraw False

    DrawControl.DrawStyle = vbDot
    DrawControl.DrawMode = vbInvert
    DrawControl.Line (griff(0).left + 4, griff(0).top)-(griff(1).left + 4, griff(1).top)
    DrawControl.Line (griff(2).left + 4, griff(2).top)-(griff(3).left + 4, griff(3).top)
    DrawControl.DrawStyle = vbSolid
    DrawControl.DrawMode = 13
End Sub

Private Sub InitArc(cObjIndex As Long)

    SuperDebug "sub/fun: InitArc"
    With ObjList(cObjIndex)
        .mPosX0 = .mLeft
        .mPosY0 = .mTop + .mHeight
        .mPosX1 = .mLeft
        .mPosY1 = .mTop
        .mPosX2 = .mLeft + .mWidth
        .mPosY2 = .mTop
        .mPosX3 = .mLeft + .mWidth
        .mPosY3 = .mTop + .mHeight
    End With

End Sub

Private Sub EditArc(cObjIndex As Long, wChange As myChange, cStep As Single)

    SuperDebug "sub/fun: EditArc"
    With ObjList(cObjIndex)

        Select Case wChange

            Case toLeft
                .mPosX0 = .mPosX0 - cStep
                .mPosX1 = .mPosX1 - cStep
                .mPosX2 = .mPosX2 - cStep
                .mPosX3 = .mPosX3 - cStep

            Case toTop
                .mPosY0 = .mPosY0 - cStep
                .mPosY1 = .mPosY1 - cStep
                .mPosY2 = .mPosY2 - cStep
                .mPosY3 = .mPosY3 - cStep

            Case toRight
                .mPosX0 = .mPosX0 + cStep
                .mPosX1 = .mPosX1 + cStep
                .mPosX2 = .mPosX2 + cStep
                .mPosX3 = .mPosX3 + cStep

            Case toBottom
                .mPosY0 = .mPosY0 + cStep
                .mPosY1 = .mPosY1 + cStep
                .mPosY2 = .mPosY2 + cStep
                .mPosY3 = .mPosY3 + cStep

            Case toWidthP
                .mPosX2 = .mPosX2 + cStep
                .mPosX3 = .mPosX3 + cStep

            Case toHeightP
                .mPosY0 = .mPosY0 + cStep
                .mPosY3 = .mPosY3 + cStep

            Case toWidthN
                .mPosX0 = .mPosX0 - cStep
                .mPosX1 = .mPosX1 - cStep

            Case toHeightN
                .mPosY1 = .mPosY1 - cStep
                .mPosY2 = .mPosY2 - cStep
        End Select

    End With

End Sub

Private Sub PrintWordWrap(cText As String, _
                          cLeft As Single, _
                          cTop As Single, _
                          cWidth As Single, _
                          cHeight As Single, _
                          Optional cAlign As AlignmentConstants = vbLeftJustify)

    SuperDebug "sub/fun: PrintWordWrap"
    Dim CrArray() As String
    Dim SpArray() As String
    Dim CrQty     As Integer
    Dim SpQty     As Integer
    Dim n1        As Integer
    Dim n2        As Integer
    Dim wQty      As Integer
    Dim wPos      As Integer
    Dim tStr      As String

    CrArray = Split(cText, vbCrLf)
    CrQty = UBound(CrArray)

    DrawControl.CurrentY = cTop

    For n1 = 0 To CrQty
        SpArray = Split(CrArray(n1), " ")
        SpQty = UBound(SpArray)
        wQty = SpQty
        wPos = 0

        Do
            Do
                tStr = ""

                For n2 = wPos To wQty
                    tStr = tStr & SpArray(n2) & " "
                Next n2

                wQty = wQty - 1
            Loop While DrawControl.TextWidth(tStr) > cWidth

            If Len(tStr) > 0 Then

                Select Case cAlign

                    Case vbCenter
                        DrawControl.CurrentX = cLeft + ((cWidth / 2) - (DrawControl.TextWidth(tStr) / 2))

                    Case vbLeftJustify
                        DrawControl.CurrentX = cLeft

                    Case vbRightJustify
                        DrawControl.CurrentX = cLeft + cWidth - DrawControl.TextWidth(tStr)
                End Select

                DrawControl.Print tStr
                wPos = wQty + 2
                wQty = SpQty
            Else
                Exit Do
            End If

            If DrawControl.CurrentY > cTop + cHeight Then Exit Do
        Loop

    Next n1

End Sub

Private Sub DrawRoundRectangle(cLeft As Single, _
                               cTop As Single, _
                               cWidth As Single, _
                               cHeight As Single, _
                               Optional cRound As Integer = 25, _
                               Optional cAngle As Single)
    SuperDebug "sub/fun: DrawRoundRectangle"
    Dim tCoor          As myCoorType
    Dim n              As Integer
    Dim POINT(1 To 25) As POINTAPI
    Dim sR             As Single
    Dim cR             As Single

    tCoor = GetCoordinate(cLeft, cTop, cWidth, cHeight, cAngle)

    sR = Sin(cAngle * Pi / 180) * cRound
    cR = Cos(cAngle * Pi / 180) * cRound

    With tCoor

        POINT(1).X = .posX2 - sR
        POINT(1).Y = .posY2 - cR

        POINT(2).X = .posX2
        POINT(2).Y = .posY2

        POINT(3).X = .posX2
        POINT(3).Y = .posY2

        POINT(4).X = .posX2 + cR
        POINT(4).Y = .posY2 - sR

        POINT(5).X = .posX2 + cR
        POINT(5).Y = .posY2 - sR

        POINT(6).X = .posX4 - cR
        POINT(6).Y = .posY4 + sR

        POINT(7).X = .posX4 - cR
        POINT(7).Y = .posY4 + sR

        POINT(8).X = .posX4
        POINT(8).Y = .posY4

        POINT(9).X = .posX4
        POINT(9).Y = .posY4

        POINT(10).X = .posX4 - sR
        POINT(10).Y = .posY4 - cR

        POINT(11).X = .posX4 - sR
        POINT(11).Y = .posY4 - cR

        POINT(12).X = .posX3 + sR
        POINT(12).Y = .posY3 + cR

        POINT(13).X = .posX3 + sR
        POINT(13).Y = .posY3 + cR

        POINT(14).X = .posX3
        POINT(14).Y = .posY3

        POINT(15).X = .posX3
        POINT(15).Y = .posY3

        POINT(16).X = .posX3 - cR
        POINT(16).Y = .posY3 + sR

        POINT(17).X = .posX3
        POINT(17).Y = .posY3

        POINT(18).X = .posX1 + cR
        POINT(18).Y = .posY1 - sR

        POINT(19).X = .posX1 + cR
        POINT(19).Y = .posY1 - sR

        POINT(20).X = .posX1
        POINT(20).Y = .posY1

        POINT(21).X = .posX1
        POINT(21).Y = .posY1

        POINT(22).X = .posX1 + sR
        POINT(22).Y = .posY1 + cR

        POINT(23).X = .posX1 + sR
        POINT(23).Y = .posY1 + cR

        POINT(24).X = .posX2 - sR
        POINT(24).Y = .posY2 - cR

        POINT(25).X = .posX2 - sR
        POINT(25).Y = .posY2 - cR

    End With

    BeginPath DrawControl.hdc
    PolyBezier DrawControl.hdc, POINT(1), 25
    EndPath DrawControl.hdc
    StrokeAndFillPath DrawControl.hdc
End Sub

Public Sub DrawGIS(PicLoad As Variant)
    SuperDebug "sub/fun: DrawGIS"
    Dim oRect As New XRect
    
    With picMap
        '.Move .Left, .Top, GIS.Width, GIS.Height
        .Width = GIS.Width
        .Height = GIS.Height
        oRect.Prepare 0, 0, PicLoad.Width, PicLoad.Height 'ScaleX(.Width, vbTwips, vbPixels), ScaleY(.Height, vbTwips, vbPixels)
        '.ZOrder 0
        '.Cls
    End With

    'PicLoad.AutoRedraw = True
    GIS.PrintDC PicLoad.hdc, 96, oRect, GIS.viewer.VisibleExtent, 0
    GIS.PrintDC picMap.hdc, 96, oRect, GIS.viewer.VisibleExtent, 0
    GIS.draw
    picMap.Picture = picMap.Image
    'Debug.Print PicData.Picture
    PicLoad.Picture = picMap.Image
    PicLoad.Visible = True
End Sub

Public Sub DrawGIS2(ileft As Integer, _
                    itop As Integer, _
                    iwidth As Integer, _
                    iheight As Integer, _
                    oPic As StdPicture, Optional bUnRedraw As Boolean)
    SuperDebug "sub/fun: DrawGIS2"
    
    If bUnRedraw Then Exit Sub
    
    Dim oRect As New XRect
    Dim oPB   As PictureBox

    picMap.Cls
    picMap.Move ileft, itop, iwidth, iheight
    oRect.Prepare 0, 0, iwidth, iheight 'ScaleX(.Width, vbTwips, vbPixels), ScaleY(.Height, vbTwips, vbPixels)
    
    GIS.PrintDC picMap.hdc, 96, oRect, GIS.viewer.VisibleExtent, 0
    Set oPic = picMap.Image
End Sub

Public Sub InitGIS(PicLoad As Variant, Optional dHeight As Double, Optional dWidth As Double)
'    Dim ll As XGIS_LayerSHP
'
'    Set ll = New XGIS_LayerSHP
'    ll.Path = GisUtils.GisSamplesDataDir + "\World\Countries\Poland\DCW\country.shp"
'    ll.Name = "states"
    SuperDebug "sub/fun: InitGIS"
    If Not dHeight = 0 And Not dWidth = 0 Then
        GIS.Move GIS.left, GIS.top, dWidth, dHeight
    End If
'    GIS.Draw
    
'    GIS.Close
'
'    GIS.Add ll
'
'    Set ll = New XGIS_LayerSHP
'    ll.Path = GisUtils.GisSamplesDataDir + "\World\Countries\Poland\DCW\lwaters.shp"
'    ll.Name = "rivers"
'    ll.UseConfig = False
'    ll.Params.Line.OutlineWidth = 0
'    ll.Params.Line.Width = 3
'    ll.Params.Line.color = RGB(0, 0, 255)
'    GIS.Add ll
'
'    GIS.FullExtent
    Legend.GIS_Viewer = GIS.viewer
    NorthArrow.GIS_Viewer = GIS.viewer
    ScaleBar.GIS_Viewer = GIS.viewer
    DrawGIS PicLoad
    
    Dim i As Long
    Dim oLayer As XGIS_LayerAbstract
    Dim oList As TatukGIS_XDK10.XList
    Set oList = GIS.viewer.items
    
    
    Do Until i = GIS.items.Count
    
        Set oLayer = GIS.viewer.get(oList.Item(i).Name)
        
        If Not oLayer.Active Then
            oLayer.HideFromLegend = True
        End If
    
    i = i + 1
    Loop
    Legend.UpDate
    
    Set oLayer = Nothing
    Set oList = Nothing
    
End Sub

Public Sub SetGISComponent(sPath As String, xmin As Double, xmax As Double, ymin As Double, ymax As Double)
        '<EhHeader>
        On Error GoTo SetGISComponent_Err
        '</EhHeader>
100 SuperDebug "sub/fun: SetGISComponent"
        Dim oExtent As New XGIS_Extent

102     GIS.Open sPath, False
    
104     With oExtent
106         .xmin = xmin
108         .ymin = ymin
110         .xmax = xmax
112         .ymax = ymax
        End With
    
114     GIS.VisibleExtent = oExtent
    
        '<EhFooter>
        Exit Sub

SetGISComponent_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISDrawObj.SetGISComponent", _
                  "OASISDrawObj component failure"
        '</EhFooter>
End Sub

Public Sub SetSelection(sLayer As String, _
                        sScope As String)
        '<EhHeader>
        On Error GoTo SetSelection_Err
        '</EhHeader>
        Dim oSHP As TatukGIS_XDK10.XGIS_Shape
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector

100     Set oLyr = GIS.get(sLayer)
    
102     For Each shp In oLyr.Loop(oLyr.Extent, sScope, Nothing, "", True)

104         If Not shp Is Nothing Then
106             Set oSHP = shp.MakeEditable
108             oSHP.IsHidden = False
                oSHP.IsSelected = True
110             oSHP.Invalidate True
112             Set oSHP = Nothing
            End If

        Next
    
114     Debug.Print sLayer
116     Debug.Print sScope
        '<EhFooter>
        Exit Sub

SetSelection_Err:
        Err.Raise vbObjectError + 100, "OASISClient.OASISDrawObj.SetSelection", "OASISDrawObj component failure"
        '</EhFooter>
End Sub

Public Sub SetMapProject(sPath As String)
    SuperDebug "sub/fun: SetMapProject"
    GIS.Open sPath, False
End Sub

Public Sub InitLegend(PicLoad As Variant)
    SuperDebug "sub/fun: InitLegend"
    DrawLegend Legend.left, Legend.top, Legend.Width, Legend.Height, PicLoad.Picture
End Sub

Public Sub InitScale(PicLoad As Variant)
    SuperDebug "sub/fun: InitScale"
    DrawScaleBar ScaleBar.left, ScaleBar.top, ScaleBar.Width, ScaleBar.Height, PicLoad.Picture
End Sub

Public Sub InitNorthArrow(PicLoad As Variant)
    SuperDebug "sub/fun: InitNorthArrow"
    DrawNorthArrow NorthArrow.left, NorthArrow.top, NorthArrow.Width, NorthArrow.Height, PicLoad.Picture
End Sub

Public Sub DrawLegend(ileft As Integer, _
                      itop As Integer, _
                      iwidth As Integer, _
                      iheight As Integer, _
                      oPic As StdPicture)
    SuperDebug "sub/fun: DrawLegend"
    'Dim oRect As New XRect
    'Dim oPB   As PictureBox

    'picMap.Cls
    'picMap.Move ileft, itop, iwidth, iheight
    'oRect.Prepare 0, 0, ScaleX(iwidth, vbPixels, vbTwips), ScaleY(iheight, vbPixels, vbTwips)
    Legend.Move Legend.left, Legend.top, iwidth, iheight
    'Legend.PrintDC picMap.hdc, 300, oRect, GIS.Viewer.ScaleAsFloat
    'Legend.PrintClipboard
    'Set oPic = Clipboard.GetData
    Legend.PrintBmp oPic, 10000, -2000, 0, GIS.ScaleAsFloat    '10000
    'Set oPic = picMap.Image
    'ExportLegend
End Sub

Public Sub DrawScaleBar(ileft As Integer, _
                      itop As Integer, _
                      iwidth As Integer, _
                      iheight As Integer, _
                      oPic As StdPicture)
    SuperDebug "sub/fun: DrawScaleBar"
    Dim oRect As New XRect
    Dim oPB   As PictureBox

    picMap.Cls
    picMap.Move ileft, itop, iwidth, iheight
    oRect.Prepare 0, 0, ScaleX(iwidth, vbPixels, vbTwips), ScaleY(iheight, vbPixels, vbTwips)
    ScaleBar.Move ScaleBar.left, ScaleBar.top, iwidth, iheight
    ScaleBar.PrintDC picMap.hdc, 96, oRect, GIS.viewer.ScaleAsFloat
    ScaleBar.PrintClipboard GIS.ScaleAsFloat
    Set oPic = Clipboard.GetData

End Sub

Public Sub DrawNorthArrow(ileft As Integer, _
                      itop As Integer, _
                      iwidth As Integer, _
                      iheight As Integer, _
                      oPic As StdPicture)
    SuperDebug "sub/fun: DrawNorthArrow"
    Dim oRect As New XRect
    Dim oPB   As PictureBox

    picMap.Cls
    picMap.Move ileft, itop, iwidth, iheight
    oRect.Prepare 0, 0, ScaleX(iwidth, vbPixels, vbTwips), ScaleY(iheight, vbPixels, vbTwips)
    NorthArrow.Move NorthArrow.left, NorthArrow.top, iwidth, iheight
    NorthArrow.PrintDC picMap.hdc, 96, oRect
    NorthArrow.PrintClipboard
    Set oPic = Clipboard.GetData

End Sub

Private Sub ExportLegend()
    
    SuperDebug "sub/fun: ExportLegend"
    Dim picStdPicture       As New StdPicture
    Dim picOutput           As New StdPicture
    
    Dim lTotalHeight        As Long
    Dim lControlHeight      As Long
    Dim dNumberOfIterations As Double
    Dim i                   As Integer
    Dim T                   As Long
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' HIDE THE UNSELECTED LAYERS
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Dim l                   As Long
    Dim b2                  As Boolean
    Dim b3                  As Boolean
    Dim oLyrAB              As TatukGIS_XDK10.XGIS_LayerAbstract
    Dim bLayerExists        As Boolean
    Dim sLayerNames()       As String
    Dim bLayerSeleted()     As Boolean
        
    l = 0
    ReDim sLayerNames(l + 1)
    ReDim bLayerSeleted(l + 1)
        
    bLayerExists = False
    On Error Resume Next
    bLayerExists = Legend.GetLayerInfo(l, sLayerNames(l), bLayerSeleted(l), b2, b3)
    On Error GoTo 0
        
    Do Until Not bLayerExists
        
        If Not bLayerSeleted(l) Then
            
            Set oLyrAB = Legend.GIS_Viewer.get(sLayerNames(l))
            oLyrAB.HideFromLegend = True

        End If
        
        l = l + 1
        ReDim Preserve sLayerNames(l + 1)
        ReDim Preserve bLayerSeleted(l + 1)
            
        bLayerExists = False
        On Error Resume Next
        bLayerExists = Legend.GetLayerInfo(l, sLayerNames(l), bLayerSeleted(l), b2, b3)
        On Error GoTo 0
    Loop
        
    Legend.UpDate
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Picture1.AutoSize = True
    Picture2.AutoSize = True
    Picture1.AutoRedraw = True
    Picture2.AutoRedraw = True
    Clipboard.Clear

    Legend.Width = Legend.Width * 2
    Legend.PrintClipboard
    
    Set picOutput = Clipboard.GetData(vbCFMetafile)
    lTotalHeight = picOutput.Height / Screen.TwipsPerPixelY
    lTotalHeight = lTotalHeight / 2.35
    
    lControlHeight = Legend.Height / Screen.TwipsPerPixelY
    
    dNumberOfIterations = lTotalHeight / lControlHeight

    Picture1.Height = picOutput.Height
    Picture1.Width = picOutput.Width
    Picture1.PaintPicture picOutput, 0, 0

    T = 0

    Do Until i > dNumberOfIterations

        Legend.PrintBmp picStdPicture, 0, lControlHeight, T, 1
        
        Picture2.Height = picStdPicture.Height
        Picture2.Width = picStdPicture.Width
        Picture2.PaintPicture picStdPicture, 0, 0
        BitBlt Picture1.hdc, 0, -T, picStdPicture.Width, picStdPicture.Height, Picture2.hdc, 0, 0, vbSrcCopy
        T = T - lControlHeight
        i = i + 1
    Loop
   
    Picture2.Cls
    Picture2.Height = lTotalHeight * Screen.TwipsPerPixelY
    Picture2.Width = Picture2.Width * 0.5
    BitBlt Picture2.hdc, 0, 0, Picture2.Width, Picture2.Height, Picture1.hdc, 0, 0, vbSrcCopy
     
    Legend.Width = Legend.Width / 2
    Legend.PrintBmp picStdPicture, 0, 0, 0, 1
    
    Clipboard.SetData Picture2.Image
    'MsgBox "The Legend has been added to your clipboard.", vbInformation, "OASIS Client Export..."
    
    Picture2.Cls
    Picture1.Cls
    Set picStdPicture = Nothing
    Set picOutput = Nothing

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' unHIDE THE UNSELECTED LAYERS
        
    l = 0

    Do Until l = UBound(sLayerNames) - 1
        
        If Not bLayerSeleted(l) Then
            
            Set oLyrAB = Legend.GIS_Viewer.get(sLayerNames(l))
            oLyrAB.HideFromLegend = False

        End If
        
        l = l + 1
    Loop
        
    Legend.UpDate
    Set oLyrAB = Nothing
    
End Sub

Property Get MapScale() As String
    MapScale = GIS.ScaleAsText
End Property

Property Get MapRotationAngle() As Variant
    MapRotationAngle = GIS.viewer.RotationAngle
End Property

Property Let MapRotationAngle(vAngle As Variant)
    GIS.viewer.RotationAngle = vAngle
End Property

Property Get MapRotationPointX() As Double
    MapRotationPointX = GIS.viewer.RotationPoint.X
End Property

Property Let MapRotationPointX(vPT As Double)
    GIS.viewer.RotationPoint.X = vPT
End Property

Property Get MapRotationPointY() As Double
    MapRotationPointY = GIS.viewer.RotationPoint.Y
End Property

Property Let MapRotationY(vPT As Double)
    GIS.viewer.RotationPoint.Y = vPT
End Property

Property Get MapProjectName() As String
    MapProjectName = GIS.ProjectName
End Property
    
Property Get MapColor() As Variant
    MapColor = GIS.viewer.color ' color
End Property

Property Get MapLayersNumbers() As Integer
    MapLayersNumbers = GIS.items.Count
End Property
    ''Gis.ProjectFile.

Property Get MapCoordSysDesc() As String
    MapCoordSysDesc = GIS.CS.Description
End Property

Property Get MapCoordSysEPSG() As String
    MapCoordSysEPSG = GIS.CS.EPSG
End Property

Property Get MapCoordSysName() As String
    MapCoordSysName = GIS.CS.FriendlyName
End Property

Property Get MapCoordPrettyWKT() As String
    MapCoordPrettyWKT = GIS.CS.PrettyWKT
End Property

Property Get MapCoordSysWKT() As String
    MapCoordSysWKT = GIS.CS.WKT
End Property

Property Get ScaleBarColor() As Variant
    ScaleBarColor = ScaleBar.color
End Property
    
Property Get ScaleBarDividers() As Variant
    ScaleBarDividers = ScaleBar.Dividers
End Property
    
Property Get ScaleBarUnitsName() As String
    ScaleBarUnitsName = ScaleBar.Units.FriendlyName
End Property

Property Get ScaleBarUnitsType() As String
    ScaleBarUnitsType = ScaleBar.Units.UnitsType
End Property

Property Get ScaleBarUnitsWKT() As String
    ScaleBarUnitsWKT = ScaleBar.Units.WKT
End Property

Property Get ScaleBarUnitsDesc() As String
    ScaleBarUnitsDesc = ScaleBar.Units.Description
End Property

Property Get ScaleBarUnitsEPSG() As String
    ScaleBarUnitsEPSG = ScaleBar.UnitsEPSG
End Property

Property Get ScaleBarUnitsSymbol() As String
    ScaleBarUnitsSymbol = ScaleBar.Units.Symbol
End Property

Property Get ScaleBarFontColor() As String
    ScaleBarFontColor = ScaleBar.FontColor
End Property

'    ScaleBar.Font

Property Get LegendColor() As String
    LegendColor = Legend.color
End Property

Property Get NAColor() As String
    NAColor = NorthArrow.color
End Property

Property Get NAColor1() As String
    NAColor1 = NorthArrow.Color1
End Property

Property Get NAColor2() As String
    NAColor2 = NorthArrow.Color2
End Property

Property Get NAFontColor() As String
    NAFontColor = NorthArrow.FontColor
End Property

Property Get NAPath() As String
    NAPath = NorthArrow.Path
End Property

Property Get NATransparent() As String
    NATransparent = NorthArrow.TRANSPARENT
End Property

Public Sub RefreshCanvas()
    SuperDebug "sub/fun: RefreshCanvas"
    DrawControl.Refresh
End Sub

Property Get CurrentTextWidth(Str As String) As Single
    CurrentTextWidth = PicData.TextWidth(Str)
End Property

Property Get MapMetaData() As String
Dim sMetadata As String
    sMetadata = sMetadata & "Scale:        " & GIS.ScaleAsText & vbCrLf
    sMetadata = sMetadata & "Center X:    " & GIS.viewer.CenterPtg.X & vbCrLf
    sMetadata = sMetadata & "Center Y:    " & GIS.viewer.CenterPtg.Y & vbCrLf
    sMetadata = sMetadata & "Zoom:         " & GIS.viewer.zoom ' & vbCrLf
    'sMetadata = sMetadata & "Date Created: " & Now()
    MapMetaData = sMetadata
End Property


