VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Begin VB.Form frmCoordPicker 
   Caption         =   "Coordinate Picker"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5580
   Icon            =   "frmCoordPicker.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6150
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   6150
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5580
      _cx             =   9843
      _cy             =   10848
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
      BorderWidth     =   1
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
      GridRows        =   2
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmCoordPicker.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin TatukGIS_DK.XGIS_ViewerWnd GIS 
         Height          =   5430
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   5550
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
         SelectionPattern=   "frmCoordPicker.frx":6891
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
      Begin VB.Frame FraMapTools 
         Caption         =   "Map Tools:"
         Height          =   660
         Left            =   15
         TabIndex        =   2
         Top             =   5475
         Width           =   5550
         Begin VB.TextBox txtMapPath 
            Height          =   285
            Left            =   4260
            TabIndex        =   8
            Top             =   270
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.CommandButton cmdMapTools 
            Height          =   315
            Index           =   3
            Left            =   1410
            Picture         =   "frmCoordPicker.frx":111A3
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Pick Coordinate"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdOpenMapFile 
            Height          =   315
            Left            =   1800
            Picture         =   "frmCoordPicker.frx":179F5
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Open Map Project"
            Top             =   240
            Width           =   405
         End
         Begin VB.CommandButton cmdMapTools 
            Height          =   315
            Index           =   2
            Left            =   1020
            Picture         =   "frmCoordPicker.frx":1E247
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdMapTools 
            Height          =   315
            Index           =   1
            Left            =   600
            Picture         =   "frmCoordPicker.frx":24A99
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   405
         End
         Begin VB.CommandButton cmdMapTools 
            Height          =   315
            Index           =   0
            Left            =   180
            Picture         =   "frmCoordPicker.frx":2B2EB
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblCoords 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2250
            TabIndex        =   9
            Top             =   120
            Width           =   1845
         End
      End
   End
   Begin TatukGIS_DK.XGIS_ControlLegend Legend1 
      Height          =   1125
      Left            =   1290
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
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
      ReverseOrder    =   0   'False
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
Attribute VB_Name = "frmCoordPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_sAdmValues() As String
Private GisUtils As New XGIS_Utils
Public Event coords(X As Double, Y As Double, z As Double)
Private m_DouxCoord As Double
Private m_DouyCoord As Double
Private m_DouZoomVal As Double
Private m_VarAdmNames As Variant
Private m_bAutoCloseMess As Boolean

Public Property Get AutoCloseMess() As Boolean
    AutoCloseMess = m_bAutoCloseMess
End Property

Public Property Let AutoCloseMess(ByVal bValue As Boolean)
    m_bAutoCloseMess = bValue
End Property

Public Property Get AdmNames() As Variant
    AdmNames = m_VarAdmNames
End Property

Public Property Get ZoomVal() As Double
    ZoomVal = m_DouZoomVal
End Property

Public Property Let ZoomVal(ByVal DouValue As Double)
    m_DouZoomVal = DouValue
End Property

Public Property Get yCoord() As Double
    yCoord = m_DouyCoord
End Property

Public Property Let yCoord(ByVal DouValue As Double)
    m_DouyCoord = DouValue
End Property

Public Property Get xCoord() As Double
    xCoord = m_DouxCoord
End Property

Public Property Let xCoord(ByVal DouValue As Double)
    m_DouxCoord = DouValue
End Property

Public Sub Init(Optional sMapProduct As String, Optional sAdmArray As Variant)
        '<EhHeader>
        On Error GoTo init_Err
        '</EhHeader>
    Dim i As Integer
        
        lblCoords.Caption = ""
        
100     If sMapProduct = "" Then
102       cmdOpenMapFile_Click
        Else
104         GIS.Open sMapProduct, False
            
        End If
    
    
        'Check for Admin Layers
106     If TestArray(sAdmArray) Then
108         ReDim m_sAdmValues(UBound(sAdmArray))
110         For i = LBound(sAdmArray) To UBound(sAdmArray)
112             m_sAdmValues(i) = sAdmArray(i)
            Next
        End If
    
        Legend1.GIS_Viewer = GIS.Viewer
    
        '<EhFooter>
        Exit Sub

init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmCoordPicker.Init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function TestArray(MyArray As Variant) As Boolean
    On Error GoTo Trap
    Dim lSize As Long
    
    lSize = LBound(MyArray)
    TestArray = True
    
Trap:
    Err.Clear
End Function

Public Sub AddLyr(oLyr As XGIS_LayerVector)
    GIS.Add oLyr
    Legend1.Update
End Sub

Public Sub RefreshLyr(sLayer As String)
    'GIS.Items(sLayer).Paint
    Legend1.Update
End Sub

Private Sub cmdMapTools_Click(Index As Integer)
    Select Case Index
    
        Case 0
            GIS.Mode = XgisZoom
        Case 1
            GIS.Mode = XgisZoomEx
        Case 2
            GIS.Mode = XgisDrag
        Case 3
            GIS.Mode = XgisUserDefined
            GIS.CursorForUserDefined = -3
    End Select
End Sub

Private Sub Form_Load()
    m_bAutoCloseMess = True
End Sub

Private Sub GIS_OnMouseUp(translated As Boolean, _
                          ByVal Button As TatukGIS_DK.XMouseButton, _
                          ByVal Shift As TatukGIS_DK.XShiftState, _
                          ByVal X As Long, _
                          ByVal Y As Long)
        '<EhHeader>
        On Error GoTo GIS_OnMouseUp_Err
        '</EhHeader>
        Dim ptg As XGIS_Point
        Dim shp As XGIS_Shape
        Dim i As Integer
  
        Select Case GIS.Mode
        
            Case XgisUserDefined
100             Set ptg = GIS.ScreenToMap(GisUtils.Point(X, Y))
        
102             m_DouxCoord = ptg.X
104             m_DouyCoord = ptg.Y
106             m_DouZoomVal = GIS.Zoom
                'm_VarAdmNames
108             RaiseEvent coords(ptg.X, ptg.Y, GIS.Zoom)
110             Me.Caption = "X:" & ptg.X & " Y:" & ptg.Y & " Zoom:" & GIS.Zoom
                lblCoords.Caption = "X:" & ptg.X & vbCrLf & "Y:" & ptg.Y & vbCrLf & "Zoom:" & GIS.Zoom
                
                If m_bAutoCloseMess Then
                    If MsgBox("Current values: " & lblCoords.Caption & vbCrLf & "If these are correct click yes.", vbYesNo, "OASIS Admin Coord picker") = vbYes Then
                        Me.Hide
                    End If
                End If
        
        End Select
        
        '<EhFooter>
        Exit Sub

GIS_OnMouseUp_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmCoordPicker.GIS_OnMouseUp " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub cmdOpenMapFile_Click()
        '<EhHeader>
        On Error GoTo cmdOpenMapFile_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog
    
        On Error Resume Next
100     c.DefaultExt = "*.ttkgp"
102     c.DialogTitle = "Open Map Definition File"
104     c.Filter = "Map Definition Files (*.ttkgp;*.prj)|*.ttkgp;*.prj"
106     c.ShowOpen
108     txtMapPath.Text = c.Filename
        GIS.Open txtMapPath.Text, False
        '<EhFooter>
        Exit Sub

cmdOpenMapFile_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in GeoMarksExplorer.frmGeoMarksExplorer.cmdOpenMapFile_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


