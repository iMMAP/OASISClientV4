VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{9C989D2F-3596-477B-B719-6DC4E6893AF2}#1.0#0"; "TATUKG~1.OCX"
Begin VB.Form frmOVMap 
   Caption         =   "Overview Map"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2760
   LinkTopic       =   "Form2"
   ScaleHeight     =   2880
   ScaleWidth      =   2760
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elMiniMapHolder 
      Height          =   2880
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2760
      _cx             =   4868
      _cy             =   5080
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
      AutoSizeChildren=   8
      BorderWidth     =   0
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
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmOVMap.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin TatukGIS_XDK10.XGIS_ViewerWnd GISm 
         Height          =   2355
         Left            =   0
         TabIndex        =   2
         Top             =   195
         Width           =   2760
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
         SelectionPattern=   "frmOVMap.frx":004B
         SelectionTransparency=   100
         SelectionWidth  =   100
         SelectionOutlineOnly=   0   'False
         OldCachedPaint  =   0   'False
         PrinterModeDraft=   0   'False
         PrinterModeForceBitmap=   0   'False
         GDIType         =   1
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
         Cursor          =   0
         DoubleBuffered  =   0   'False
         ModeMouseButton =   0
         CursorForUserDefined=   0
         View3D          =   0   'False
      End
      Begin ActiveBar3LibraryCtl.ActiveBar3 abStatus 
         Height          =   300
         Left            =   0
         TabIndex        =   1
         Top             =   2580
         Width           =   2760
         _LayoutVersion  =   2
         _ExtentX        =   4868
         _ExtentY        =   529
         _DataPath       =   ""
         Bands           =   "frmOVMap.frx":1CFD
      End
   End
End
Attribute VB_Name = "frmOVMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private GIS As TatukGIS_XDK10.XGIS_Viewer
Public Event MapOnMouseDown(translated As Boolean, ByVal Button As TatukGIS_XDK10.XMouseButton, ByVal Shift As TatukGIS_XDK10.XShiftState, ByVal x As Long, ByVal y As Long)
Public Event MapOnMouseMove(translated As Boolean, ByVal Shift As TatukGIS_XDK10.XShiftState, ByVal x As Long, ByVal y As Long)
Public Event MapOnMouseUp(translated As Boolean, ByVal Button As TatukGIS_XDK10.XMouseButton, ByVal Shift As TatukGIS_XDK10.XShiftState, ByVal x As Long, ByVal y As Long)

Private Sub Form_Load()
If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
End Sub

Public Sub Init(oGIS As TatukGIS_XDK10.XGIS_Viewer)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        Dim lv As TatukGIS_XDK10.XGIS_LayerVector
        Dim lW As TatukGIS_XDK10.XGIS_LayerVector

        Set GIS = oGIS
        
100     Set lv = New TatukGIS_XDK10.XGIS_LayerVector 'minimap transparent rectangle
102     lv.Transparency = 30
104     lv.Params.area.Color = vbRed
106     lv.Params.area.OutlineWidth = -2
108     lv.Name = MINIMAP_R_NAME
110     Me.GISm.Add lv

112     Set minishp = Me.GISm.get(MINIMAP_R_NAME).CreateShape(XgisShapeTypePolygon)
114     Set lW = New TatukGIS_XDK10.XGIS_LayerVector
116     lW.Params.Line.Color = RGB(0, 0, &H80&)
118     lW.Params.Line.Width = -2
120     lW.Name = MINIMAP_O_NAME
122     Me.GISm.Add lW

124     Set minishpo = Me.GISm.get(MINIMAP_O_NAME).CreateShape(XgisShapeTypeArc)

126     Me.GISm.FullExtent

128     Me.GISm.RestrictedExtent = Me.GISm.Extent
130     minishp.layer.Extent = Me.GISm.Extent
        
132     If Not Me.Visible Then
134         Me.Show vbModeless, Me
        End If
        
136     fminiMove = False

        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmOVMap.Init " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GISm_OnMouseDown(translated As Boolean, ByVal Button As TatukGIS_XDK10.XMouseButton, ByVal Shift As TatukGIS_XDK10.XShiftState, ByVal x As Long, ByVal y As Long)
        '<EhHeader>
        On Error GoTo GISm_OnMouseDown_Err
        '</EhHeader>
100         translated = True
102         If (Button = XmbRight) Then Exit Sub
104         fminiMove = True
        '<EhFooter>
        Exit Sub

GISm_OnMouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOVMap.GISm_OnMouseDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GISm_OnMouseMove(translated As Boolean, ByVal Shift As TatukGIS_XDK10.XShiftState, ByVal x As Long, ByVal y As Long)
        '<EhHeader>
        On Error GoTo GISm_OnMouseMove_Err
        '</EhHeader>
            Dim ptg As TatukGIS_XDK10.XGIS_Point
100         translated = True
102         If ((Not fminiMove) And (Not (Shift = XssCtrl))) Then Exit Sub
104         Set ptg = GISm.ScreenToMap(GisUtils.Point(x, y))
106         minishp.SetPosition miniRecalc(ptg), GISm.get(MINIMAP_R_NAME), 5
108         GISm.UpDate
110         If (Shift = XssCtrl) Then
112           RaiseEvent MapOnMouseMove(translated, Shift, x, y)
            End If
        '<EhFooter>
        Exit Sub

GISm_OnMouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOVMap.GISm_OnMouseMove " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GISm_OnMouseUp(translated As Boolean, ByVal Button As TatukGIS_XDK10.XMouseButton, ByVal Shift As TatukGIS_XDK10.XShiftState, ByVal x As Long, ByVal y As Long)
        '<EhHeader>
        On Error GoTo GISm_OnMouseUp_Err
        '</EhHeader>
100   translated = True
102   RaiseEvent MapOnMouseUp(translated, Button, Shift, x, y)
        '<EhFooter>
        Exit Sub

GISm_OnMouseUp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOVMap.GISm_OnMouseUp " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


