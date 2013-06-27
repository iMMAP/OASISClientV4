VERSION 5.00
Object = "{06C4C9AB-574E-4612-86FB-2C144A8B8F9B}#9.0#0"; "TatukGIS_XDK9.ocx"
Begin VB.Form frmSpatialAnalysisLegend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spatial Analysis Legend"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4545
   Icon            =   "frmSpatialAnalysisLegend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin TatukGIS_XDK9.XGIS_ControlLegend Legend1 
      Height          =   5190
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   4470
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
      ParentColor     =   -1  'True
      ParentCtl3D     =   -1  'True
      Object.Visible         =   -1  'True
      DoubleBuffered  =   -1  'True
      AllowMove       =   -1  'True
      AllowActive     =   -1  'True
      AllowExpand     =   -1  'True
      AllowParams     =   -1  'True
      Mode            =   0
   End
   Begin TatukGIS_XDK9.XGIS_ViewerWnd GIS 
      Height          =   1545
      Left            =   1530
      TabIndex        =   1
      Top             =   1890
      Width           =   1500
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
      SelectionPattern=   "frmSpatialAnalysisLegend.frx":6852
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
      ParentColor     =   0   'False
      ParentCtl3D     =   0   'False
      Object.Visible         =   -1  'True
      Cursor          =   16
      DoubleBuffered  =   0   'False
      ModeMouseButton =   0
      CursorForUserDefined=   0
   End
End
Attribute VB_Name = "frmSpatialAnalysisLegend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
End Sub
