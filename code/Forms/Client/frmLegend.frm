VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Begin VB.Form frmLegend 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Theme Legend"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   3720
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TatukGIS_DK.XGIS_ViewerWnd moGIS 
      Height          =   1500
      Left            =   1125
      TabIndex        =   2
      Top             =   1980
      Visible         =   0   'False
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
      SelectionPattern=   "frmLegend.frx":0000
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
   Begin C1SizerLibCtl.C1Elastic elLegend 
      Height          =   5430
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3720
      _cx             =   6562
      _cy             =   9578
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
      _GridInfo       =   $"frmLegend.frx":00CA
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin TatukGIS_DK.XGIS_ControlLegend Legend1 
         Height          =   5430
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3720
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
      End
   End
End
Attribute VB_Name = "frmLegend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Init(oGIS As Object, oLayer As TatukGIS_XDK9.XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
    Dim strName As String

100     If moGIS.Items.Count > 0 Then
102         strName = moGIS.Items.Item(0).Name
104         moGIS.Delete strName
106         Legend1.UpDate
108         moGIS.UpDate
        
        End If
    
    
    
110     moGIS.Add oLayer
    
112     Legend1.AllowActive = False
114     Legend1.AllowMove = False
116     Legend1.AllowParams = False
    
118     Legend1.GIS_Viewer = moGIS.viewer
120     Legend1.UpDate
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLegend.init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub UpDate()
    On Error Resume Next
    Legend1.UpDate
End Sub

Private Sub Test()
    
End Sub

Private Sub Form_Load()
    If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If

End Sub
