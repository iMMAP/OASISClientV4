VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SetParent _
                Lib "user32" (ByVal hWndChild As Long, _
                              ByVal hWndNewParent As Long) As Long

Public Enum OChartScheme
    CS_Solid = 0
    CS_BWPattern = 1
    CS_Pattern = 2
    CS_Gradient = 3
End Enum

Public Enum oChartAccess
    CA_None = 0
    CA_Frame3D = 1
    CA_Math = 2
    CA_FlatFrame = 3
End Enum

Public Enum oChartPalettes
    Windows = 0
    Sky = 1
    Earth_Tones = 2
    Modern_Business = 3
    Pastels = 4
    Mese = 5
    Adventure = 6
    ChartDef5 = 7
    High_Contrast = 8
    ChartDef6 = 9
    Alternate = 10
    Vivid = 11
    Alt_Pastels = 12
End Enum

Public Enum oChartBorderEffect
    CBE_None = 0
    CBE_Raised = 1
    CBE_Light = 2
    CBE_Dark = 3
    CBE_Opposite = 4
    CBE_Shadow = 5
End Enum

Public Enum oChartGrid
    CG_None = 0
    CG_Horz = 1
    CG_Vert = 2
End Enum

Public Enum oChartMarkerShape
    CMS_None = 0
    CMS_Rect = 1
    CMS_Circle = 2
    CMS_Triangle = 3
    CMS_Diamond = 4
    CMS_Marble = 5
    CMS_HorzLine = 6
    CMS_VertLine = 7
    CMS_Cross = 8
    CMS_InvertedTriangle = 9
    CMS_X = 10
    CMS_Picture = 11
    CMS_Many = 12
                
End Enum

Public Enum oChartSmoothFlag
    CSF_None = 0
    CSF_Fill = 1
    CSF_Border = 2
    CSF_Text = 4
    CSF_TextSystem = 8
End Enum

Public Enum oChartStacked
    CS_No = 0
    CS_Normal = 1
    CS_Stacked100 = 2
End Enum

Public Enum oChartExportFormat
    imgBMP = FileFormat.FileFormat_Bitmap
    imgWMF = FileFormat.FileFormat_Metafile
    tplBin = FileFormat.FileFormat_Binary
    tplXML = FileFormat.FileFormat_XML
    dmpData = FileFormat.FileFormat_BinaryTemplate
End Enum

Public Type oChartExports
    sPath As String
    sFileName As String
    enmExportFormat As oChartExportFormat
    bForceKill As Boolean
End Type

Public Enum OASISChartType
    Chart_Area = 3
    Chart_Bar = 2
    Chart_Bubble = 15
    Chart_Candlestick = 17
    Chart_Contour = 18
    Chart_Cube = 12
    Chart_Curve = 6
    Chart_CurveArea = 19
    Chart_Doughnut = 13
    Chart_Gantt = 20
    Chart_HiLowClose = 9
    Chart_Lines = 1
    Chart_OpenHiLowClose = 16
    Chart_Pareto = 7
    Chart_Pie = 5
    Chart_Pyramid = 14
    Chart_Radar = 11
    Chart_Scatter = 4
    Chart_Step = 8
    Chart_Surface = 10
End Enum

Public Type OASISChartGeneric
    bUseAppConnection As Boolean
    sName As String
    sDescription As String
    lParentHwnd As Long
End Type

Public Enum oChartAlignment
    CA_StringAlignment_Center = 1
    CA_StringAlignment_Far = 2
    CA_StringAlignment_Near = 0
    CA_Docked_Left = 513
    CA_Docked_Right = 515
    CA_Docked_Top = 256
    CA_Docked_Bottom = 258
End Enum

Public Enum OChartArea
    CDA_Area_Bottom = 3
    CDA_Area_Left = 0
    CDA_Area_Right = 2
    CDA_Area_Top = 1

End Enum

Public Enum oChartFlags
    CF_TitleFlag_DrawingArea = 1
    CF_TitleFlag_ExcludeFromBorder = 4
    CF_TitleFlag_FixInverted = 16
    CF_TitleFlag_Hidden = 2
    CF_TitleFlag_HiddenByBorder = 32678
    CF_TitleFlag_NoChangeDockArea = 8
End Enum

Public Type OChartFonts
    CF_Bold As Boolean
    CF_Italic As Boolean
    CF_Name As String
    CF_Size As Integer
    CF_Strikethrough As Boolean
    CF_Underline As Boolean
    CF_Weight As Integer
End Type

Public Type oChartTitles
    CT_Alignment As oChartAlignment
    CT_BackColor  As Long
    CT_DockArea As OChartArea
    CT_DrawingArea As Boolean
    CT_Flags As oChartFlags
    CT_Font As OChartFonts
    CT_Gap As Integer
    CT_LineAlignment As oChartAlignment
    CT_LineGap As Integer
    CT_Link As String
    CT_Text As String
    CT_TextColor As Long
    CT_Url As String
End Type

Public Type OChartTemplate
    sName As String
    sDecription As String
    enmFormat As oChartExportFormat
End Type

Public Enum Y_AxisFormat
    Y_AxisFormat_None = 0
    Y_AxisFormat_Number = 1
    Y_AxisFormat_Currency = 2
    Y_AxisFormat_Scientific = 3
    Y_AxisFormat_Percentage = 4
    Y_AxisFormat_Date = 5
    Y_AxisFormat_LongDate = 6
    Y_AxisFormat_Time = 7
    Y_AxisFormat_DateTime = 8
End Enum

Public Type OASISChartObj
    udtGeneric As OASISChartGeneric
    sConnStr As String
    sSQL As String
    bChartTBR As Boolean
    bAnnoTBR As Boolean
    bDataEdtr As Boolean
    bSeriesLGD As Boolean
    bValueLGD As Boolean
    bDataHigls As Boolean
    sChartTools() As String
    sAnnoTools() As String
    udtTitle As oChartTitles
    udtNotes As oChartTitles
    sXAxis As String
    sYAxis As String
    bContextMenu As Boolean
    bMenuBar As Boolean
    bDataEdtrAllowEdit As Boolean
    bDataEdtrAllowDrag As Boolean
    bMultipleColors As Boolean
    bPointLabelsGen As Boolean
    bScrollable As Boolean
    bHGlsDimmed As Boolean
    bHGlsPointLabel As Boolean
    iSerLgdAlign As Integer
    iValLgdAlign As Integer
    iDataEdtrAlign As Integer
    iWidth As Integer
    iHeight As Integer
    iParentWidth As Integer
    iParentHeight As Integer
    enmChartType As OASISChartType
    udtExports(4) As oChartExports
    udtChartTemplate As OChartTemplate
    enmScheme As OChartScheme
    iAngleX As Integer
    iAngleY As Integer
    enmAxesStyle As oChartAccess
    lngBackColor As Long
    bBorder As Boolean
    lngBorderColor As Long
    enmBorderEffect As oChartBorderEffect
    bCluster As Boolean
    bChart3D As Boolean
    bCrossHairs As Boolean
    sngCylSides As Single ' FROM -32 TO 32
    enmGrid As oChartGrid
    lngInsideColor As Long
    enmMarkerShape As oChartMarkerShape
    iMarkerSize As Integer
    sngMarkerStep As Single
    sngPerspective As Single
    bShowTips As Boolean
    enmSmoothFlags As oChartSmoothFlag
    enmStacked As oChartStacked
    iWallWidth As Integer
    bZoom As Boolean
    enmPalette As oChartPalettes
    XAxisAngle As Single
    XAxisStaggered As Boolean
    YAxisLabelDecimals As Integer
    YAxisFormat As Y_AxisFormat
End Type

Private frmOCharting As frmChartWin
Private m_bAllowAutoResize As Boolean
Private m_bAllowAutoCloseDBLClick As Boolean

Public Property Get AllowAutoCloseDBLClick() As Boolean
    AllowAutoCloseDBLClick = m_bAllowAutoCloseDBLClick
End Property

Public Property Let AllowAutoCloseDBLClick(ByVal bValue As Boolean)
    m_bAllowAutoCloseDBLClick = bValue
End Property

Public Property Get AllowAutoResize() As Boolean
    AllowAutoResize = m_bAllowAutoResize
End Property

Public Property Let AllowAutoResize(ByVal bValue As Boolean)
    m_bAllowAutoResize = bValue
End Property

Public Function GetOASISObjSettings(sTemplate As String) As OASISChartObj
        '<EhHeader>
        On Error GoTo GetOASISObjSettings_Err
        '</EhHeader>
        Dim objOASISChart As OASISChartObj
    
100     If frmOCharting Is Nothing Then
102         Set frmOCharting = New frmChartWin
        End If
    
104     With frmOCharting
106         objOASISChart = .GetOASISChartObj(sTemplate)
108         .cmdExport.Visible = False
110         .enmRequestType = Init
112         .AllowAutoResize = m_bAllowAutoResize
114         .AllowAutoCloseDBLClick = m_bAllowAutoCloseDBLClick
116         .Show
118         .Width = objOASISChart.iWidth
120         .Height = objOASISChart.iHeight
122         objOASISChart.udtGeneric.lParentHwnd = .hwnd
        End With
    
124     GetOASISObjSettings = objOASISChart
        '<EhFooter>
        Exit Function

GetOASISObjSettings_Err:
        Err.Raise vbObjectError + 100, "OASISCharting.ChartProvider.GetOASISObjSettings", "ChartProvider component failure"
        '</EhFooter>
End Function

Public Function InitChart(hwnd As Long, lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long) As Long
        '<EhHeader>
        On Error GoTo InitChart_Err
        '</EhHeader>
    
100     If frmOCharting Is Nothing Then
102         Set frmOCharting = New frmChartWin
        End If
    
104     With frmOCharting
106       '  .enmRequestType = Init
108        ' .m_bAutoExport = bAutoExport
110        ' .FilePAth = sFilePath
112         '.Move Screen.Width / 2 + .Width / 2, Screen.Height / 2, 2000, 2000
'.Move 0, 0
114        ' .AllowAutoResize = m_bAllowAutoResize
116        ' .AllowAutoCloseDBLClick = m_bAllowAutoCloseDBLClick
        
118       '  If Not bAutoExport Then
120        '     .Width = udtOASISChart.iWidth
122        '     .Height = udtOASISChart.iHeight
124             .BorderStyle = vbSizable
           ' End If
    
126       '  If Not oContainer Is Nothing Then
    
128         '    Set .oContainer = oContainer
        
           ' End If
            
            InitChart = .hwnd
            
130       '  .setOASISChartObj udtOASISChart
    
132         If hwnd > 0 Then
134             SetParent .hwnd, hwnd
136             .Top = lTop
138             .Left = lLeft
                .Width = lWidth
                .Height = lHeight
                
142             .Show
                .Move 0, 0
            Else
            
               ' .Show vbModal
            End If
    
146
        End With

        '<EhFooter>
        Exit Function

InitChart_Err:
        Err.Raise vbObjectError + 100, "OASISCharting.ChartProvider.InitChart", "ChartProvider component failure"
        '</EhFooter>
End Function

Public Sub CloseObj()
    On Error Resume Next

    If Not frmOCharting Is Nothing Then
        Unload frmOCharting
    End If

End Sub

Public Sub UpdateDataSet(RS As ADODB.Recordset)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.DataSource = RS
    End If

End Sub

Public Sub ExportChart(enmFormat As oChartExportFormat, _
                       FilePAth As Variant)
    frmOCharting.Chart1.Export enmFormat, FilePAth
End Sub

Public Sub ShowChart(bVisible, _
                     Optional lParentHwnd As Long)
    frmOCharting.Visible = bVisible
    
    If Not lParentHwnd = 0 Then
        SetParent frmOCharting.hwnd, lParentHwnd
    End If

End Sub

Private Sub Class_Initialize()
    Set frmOCharting = New frmChartWin
End Sub

'-----------------------------------------------------
Public Property Get bChartTBR() As Boolean

    If Not frmOCharting Is Nothing Then
        bChartTBR = frmOCharting.Chart1.ToolBar
    End If

End Property

Public Property Let bChartTBR(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.ToolBar = oVal
    End If

End Property

Public Property Get bDataEdtr() As Boolean

    If Not frmOCharting Is Nothing Then
        bDataEdtr = frmOCharting.Chart1.DataEditor
    End If

End Property

Public Property Let bDataEdtr(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.DataEditor = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bDataHigls() As Boolean

    If Not frmOCharting Is Nothing Then
        bDataHigls = frmOCharting.Chart1.Highlight
    End If

End Property

Public Property Let bDataHigls(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Highlight = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get udtTitle() As String

    If Not frmOCharting Is Nothing Then
        udtTitle = frmOCharting.Chart1.Titles(0).Text
    End If

End Property

Public Property Let udtTitle(ByVal oVal As String)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Titles(0).Text = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get udtNotes() As String

    If Not frmOCharting Is Nothing Then
        udtNotes = frmOCharting.Chart1.Titles(1).Text
    End If

End Property

Public Property Let udtNotes(ByVal oVal As String)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Titles(1).Text = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get sXAxis() As String

    If Not frmOCharting Is Nothing Then
        sXAxis = frmOCharting.Chart1.XAxis
    End If

End Property

Public Property Let sXAxis(ByVal oVal As String)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.XAxis = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get sYAxis() As String

    If Not frmOCharting Is Nothing Then
        sYAxis = frmOCharting.Chart1.YAxis
    End If

End Property

Public Property Let sYAxis(ByVal oVal As String)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.YAxis = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bContextMenu() As Boolean

    If Not frmOCharting Is Nothing Then
        bContextMenu = frmOCharting.Chart1.ContextMenu
    End If

End Property

Public Property Let bContextMenu(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.ContextMenu = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bMenuBar() As Boolean

    If Not frmOCharting Is Nothing Then
        bMenuBar = frmOCharting.Chart1.MenuBar
    End If

End Property

Public Property Let bMenuBar(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.MenuBar = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bDataEdtrAllowEdit() As Boolean

    If Not frmOCharting Is Nothing Then
        bDataEdtrAllowEdit = frmOCharting.Chart1.DataEdtrAllowEdit
    End If

End Property

Public Property Let bDataEdtrAllowEdit(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.DataEdtrAllowEdit = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bDataEdtrAllowDrag() As Boolean

    If Not frmOCharting Is Nothing Then
        bDataEdtrAllowDrag = frmOCharting.Chart1.DataEdtrAllowDrag
    End If

End Property

Public Property Let bDataEdtrAllowDrag(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.DataEdtrAllowDrag = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bMultipleColors() As Boolean

    If Not frmOCharting Is Nothing Then
        bMultipleColors = frmOCharting.Chart1.MultipleColors
    End If

End Property

Public Property Let bMultipleColors(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.MultipleColors = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bPointLabelsGen() As Boolean

    If Not frmOCharting Is Nothing Then
        bPointLabelsGen = frmOCharting.Chart1.PointLabelsGen
    End If

End Property

Public Property Let bPointLabelsGen(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.PointLabelsGen = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bScrollable() As Boolean

    If Not frmOCharting Is Nothing Then
        bScrollable = frmOCharting.Chart1.Scrollable
    End If

End Property

Public Property Let bScrollable(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Scrollable = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bHGlsDimmed() As Boolean

    If Not frmOCharting Is Nothing Then
        bHGlsDimmed = frmOCharting.Chart1.HGlsDimmed
    End If

End Property

Public Property Let bHGlsDimmed(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.HGlsDimmed = oVal
    End If

End Property

'-----------------------------------------------------
'----------------------------------------------------
Public Property Get bHGlsPointLabel() As Boolean

    If Not frmOCharting Is Nothing Then
        bHGlsPointLabel = frmOCharting.Chart1.HGlsPointLabel
    End If

End Property

Public Property Let bHGlsPointLabel(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.HGlsPointLabel = oVal
    End If

End Property

Public Property Get iWidth() As Integer

    If Not frmOCharting Is Nothing Then
        iWidth = frmOCharting.Chart1.Width
    End If

End Property

Public Property Let iWidth(ByVal oVal As Integer)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Width = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get iHeight() As Integer

    If Not frmOCharting Is Nothing Then
        iHeight = frmOCharting.Chart1.Height
    End If

End Property

Public Property Let iHeight(ByVal oVal As Integer)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Height = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get iParentWidth() As Integer

    If Not frmOCharting Is Nothing Then
        iParentWidth = frmOCharting.Width
    End If

End Property

Public Property Let iParentWidth(ByVal oVal As Integer)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Width = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get iParentHeight() As Integer

    If Not frmOCharting Is Nothing Then
        iParentHeight = frmOCharting.Height
    End If

End Property

Public Property Let iParentHeight(ByVal oVal As Integer)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Height = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get enmChartType() As OASISChartType

    If Not frmOCharting Is Nothing Then
        enmChartType = frmOCharting.Chart1.ChartType
    End If

End Property

Public Property Let enmChartType(ByVal oVal As OASISChartType)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.ChartType = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------

Public Property Get enmScheme() As OChartScheme

    If Not frmOCharting Is Nothing Then
        enmScheme = frmOCharting.Chart1.Scheme
    End If

End Property

Public Property Let enmScheme(ByVal oVal As OChartScheme)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Scheme = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get iAngleX() As Integer

    If Not frmOCharting Is Nothing Then
        iAngleX = frmOCharting.Chart1.AngleX
    End If

End Property

Public Property Let iAngleX(ByVal oVal As Integer)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.AngleX = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get iAngleY() As Integer

    If Not frmOCharting Is Nothing Then
        iAngleY = frmOCharting.Chart1.AngleY
    End If

End Property

Public Property Let iAngleY(ByVal oVal As Integer)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.AngleY = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get enmAxesStyle() As oChartAccess

    If Not frmOCharting Is Nothing Then
        enmAxesStyle = frmOCharting.Chart1.AxesStyle
    End If

End Property

Public Property Let enmAxesStyle(ByVal oVal As oChartAccess)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.AxesStyle = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get lngBackColor() As Long

    If Not frmOCharting Is Nothing Then
        lngBackColor = frmOCharting.Chart1.BackColor
    End If

End Property

Public Property Let lngBackColor(ByVal oVal As Long)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.BackColor = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bBorder() As Boolean

    If Not frmOCharting Is Nothing Then
        bBorder = frmOCharting.Chart1.Border
    End If

End Property

Public Property Let bBorder(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Border = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get lngBorderColor() As Long

    If Not frmOCharting Is Nothing Then
        lngBorderColor = frmOCharting.Chart1.BorderColor
    End If

End Property

Public Property Let lngBorderColor(ByVal oVal As Long)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.BorderColor = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get enmBorderEffect() As oChartBorderEffect

    If Not frmOCharting Is Nothing Then
        enmBorderEffect = frmOCharting.Chart1.BorderEffect
    End If

End Property

Public Property Let enmBorderEffect(ByVal oVal As oChartBorderEffect)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.BorderEffect = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bCluster() As Boolean

    If Not frmOCharting Is Nothing Then
        bCluster = frmOCharting.Chart1.Cluster
    End If

End Property

Public Property Let bCluster(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Cluster = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bChart3D() As Boolean

    If Not frmOCharting Is Nothing Then
        bChart3D = frmOCharting.Chart1.Chart3D
    End If

End Property

Public Property Let bChart3D(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Chart3D = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bCrossHairs() As Boolean

    If Not frmOCharting Is Nothing Then
        bCrossHairs = frmOCharting.Chart1.CrossHairs
    End If

End Property

Public Property Let bCrossHairs(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.CrossHairs = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get sngCylSides() As Single

    If Not frmOCharting Is Nothing Then
        sngCylSides = frmOCharting.Chart1.CylSides
    End If

End Property

Public Property Let sngCylSides(ByVal oVal As Single)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.CylSides = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get enmGrid() As oChartGrid

    If Not frmOCharting Is Nothing Then
        enmGrid = frmOCharting.Chart1.Grid
    End If

End Property

Public Property Let enmGrid(ByVal oVal As oChartGrid)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Grid = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get lngInsideColor() As Long

    If Not frmOCharting Is Nothing Then
        lngInsideColor = frmOCharting.Chart1.InsideColor
    End If

End Property

Public Property Let lngInsideColor(ByVal oVal As Long)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.InsideColor = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get enmMarkerShape() As oChartMarkerShape

    If Not frmOCharting Is Nothing Then
        enmMarkerShape = frmOCharting.Chart1.MarkerShape
    End If

End Property

Public Property Let enmMarkerShape(ByVal oVal As oChartMarkerShape)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.MarkerShape = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get iMarkerSize() As Integer

    If Not frmOCharting Is Nothing Then
        iMarkerSize = frmOCharting.Chart1.MarkerSize
    End If

End Property

Public Property Let iMarkerSize(ByVal oVal As Integer)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.MarkerSize = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get sngMarkerStep() As Single

    If Not frmOCharting Is Nothing Then
        sngMarkerStep = frmOCharting.Chart1.MarkerStep
    End If

End Property

Public Property Let sngMarkerStep(ByVal oVal As Single)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.MarkerStep = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get sngPerspective() As Single

    If Not frmOCharting Is Nothing Then
        sngPerspective = frmOCharting.Chart1.Perspective
    End If

End Property

Public Property Let sngPerspective(ByVal oVal As Single)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Perspective = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bShowTips() As Boolean

    If Not frmOCharting Is Nothing Then
        bShowTips = frmOCharting.Chart1.ShowTips
    End If

End Property

Public Property Let bShowTips(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.ShowTips = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get enmSmoothFlags() As oChartSmoothFlag

    If Not frmOCharting Is Nothing Then
        enmSmoothFlags = frmOCharting.Chart1.SmoothFlags
    End If

End Property

Public Property Let enmSmoothFlags(ByVal oVal As oChartSmoothFlag)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.SmoothFlags = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get enmStacked() As oChartStacked

    If Not frmOCharting Is Nothing Then
        enmStacked = frmOCharting.Chart1.Stacked
    End If

End Property

Public Property Let enmStacked(ByVal oVal As oChartStacked)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Stacked = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get iWallWidth() As Integer

    If Not frmOCharting Is Nothing Then
        iWallWidth = frmOCharting.Chart1.WallWidth
    End If

End Property

Public Property Let iWallWidth(ByVal oVal As Integer)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.WallWidth = oVal
    End If

End Property

'-----------------------------------------------------
'-----------------------------------------------------
Public Property Get bZoom() As Boolean

    If Not frmOCharting Is Nothing Then
        bZoom = frmOCharting.Chart1.Zoom
    End If

End Property

Public Property Let bZoom(ByVal oVal As Boolean)

    If Not frmOCharting Is Nothing Then
        frmOCharting.Chart1.Zoom = oVal
    End If

End Property

''-----------------------------------------------------
''-----------------------------------------------------
'Public Property Get enmPalette() As oChartPalettes
'   If Not frmOCharting Is Nothing Then
'       enmPalette = frmOCharting.Chart1.Palette
'   End If
'End Property
'Public Property Let enmPalette(ByVal oVal As oChartPalettes)
'   If Not frmOCharting Is Nothing Then
'        frmOCharting.Chart1.enmPalette = oVal
'   End If
'End Property

'-----------------------------------------------------
'-----------------------------------------------------

Public Function ShowTheChart(hwnd As Long) As Long

    ShowTheChart = frmChartWin.ShowTheChart(hwnd)

End Function

Public Sub TemplateLoad(sPath As String)

    frmChartWin.TemplateLoad sPath
    
End Sub

Public Sub TemplateSave(sPath As String)

    frmChartWin.TemplateSave sPath
End Sub
 
 
 Public Sub getprops()
 
    frmChartWin.getprops

 End Sub
