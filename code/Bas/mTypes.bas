Attribute VB_Name = "mTypes"
Public Type IncidentLayerSettings
    CachedPaint As Boolean
    IgnoreShapeParams As Boolean
    IncrementalPaint As Boolean
    UseConfig As Boolean
    UseFileParams As Boolean
    ConfigFilePAth As String
    HideFromLegend As Boolean
    VisibleFromStart As Boolean
End Type

Public Type MapSettings
    AlwaysSaveMapStateOnExit As Boolean
    AutoScroll As Boolean
    StoreLayerParamsInProject As Boolean
    iMAPUnits As Integer
    MapRotation As Integer
    ScrollBars As Integer
    ShowWaterMark As Boolean
    ShowNorthArrow As Boolean
    NorthArrowCustomType As Boolean
    NorthArrowTransparent As Boolean
    NorthArrowType As Integer
    NorthArrowColor As Long
End Type

Public Type AlertScrollerSettings
    BackColor As Long
    frontColor As Long
    Speed As Integer
    Height As Integer
    FontSize As Single
    FontName As String
    FontBold As Boolean
    FontItalic As Boolean
    Enabled As Boolean
    ShowDuringSynch As Boolean
End Type

Public Type UrlLayerSettings
    UseExtendedInfoWin As Boolean
    AutoShutWin As Boolean
    AutoShutTime As Long
    WinWidth As Long
    WinHeight As Long
End Type

Public Type SelectionStyle
    OutLineOnly As Boolean
    Transparency As Integer
    color As Long
    Width As Integer
End Type

Public Type ExcludeType
    sTableName As String
    sFieldName As String
End Type

Public Type SynchUpdateOptions
    ManualSynchronisation As Boolean
    lMethod As Long '0 = Internet BAtch 1 = Internet Single 2 = Folder Reader
    ApplicationSettings As Boolean
    GISAttributeSettings As Boolean
    SynchLayersSettings As Boolean
    GeoMarks As Boolean
    PrintTemplates As Boolean
    MapProducts As Boolean
    AutoUpdate As Boolean
    Charts As Boolean
    Thematics As Boolean

    ForceZero As Boolean
    Feeds As Boolean
    DynamDataDefs As Boolean
End Type

Public Type ZoomToSettings
    UseMultiple As Boolean
    SaveOnExit As Boolean
End Type

Public Type LocatorSettings
    Level1 As String
    Level2 As String
End Type

Public Type ShpProps
    uID As Long
    sLayerName As String
    UseInReport As Boolean
    Editable As Boolean
    sTabCaption As String
    lCurrTab As Integer
    bSelect As Boolean
    bFlash As Boolean
    sLayerCaption As String
End Type

Public Type MapTipSetting
    Enabled As Boolean
    TipDelay As Single
    MapTipLayer As String
    MapTipField As String
    TipColor As Long
    TextColor As Long
    TipBorder As Boolean
End Type

Public Type SelectorSettings
    sSpatialOperator As String
    dBuffeLevel As Double
    bAutoZoom As Boolean
    bAutoSelect As Boolean
    bAutoFlash As Boolean
    bAutoClear As Boolean
    bEdit As Boolean
End Type

Public Type CoordTransSettings
    Semi_Major_Axis As Double
    Inverse_Flattening As Double
    Sphere As Boolean
    False_Northing As Double
    False_Easting As Double
    Lat_Of_Origin As Double
    Long_Of_Origin As Double
    Zone As Integer
End Type

Public Type MapObjectsSettings
     UseNorthArrow As Boolean
     NorthArrowType As Integer
     NorthArrowColor As Long
     NorthArrowPicture As String
     NorthArrowTransparency As Boolean
     UseScaleBar As Boolean
     UseWaterMark As Boolean
     WaterMarkPath As String
End Type

Public Type udtCoord
    lat As Double
    lon As Double
    MGRS As String
End Type

Public Enum OASISThemes
    IncidentByType = 0
    IncidentByTarget = 1
    IncidentByTimeOfDay = 2
    IncidentByReviewed = 3
End Enum

Public Enum OASISLocationType
    Temporary = 0
    Permanent = 1
    Dynamic = 2
End Enum

Public Enum OASISFeatureTypes
    Incident = 0
    Location = 1
    GeoMark = 2
    Custom = 3
    
    Selection = 5
End Enum

Public Enum OASISMenuButton
    Incidentwizard = 0
    Locationwizard = 1
    Activitieswizard = 2
    Personellwizard = 3
    Radioroom = 4
    IncidentAnalysis = 5
    LocationAnalysis = 6
End Enum

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public ptPopUpPos As POINTAPI

Public Enum OASIS_TOOLS
    oZoom = 0
    oZoomEx = 1
    oZoomIn = 2
    oZoomOut = 3
    oZoom2FullExtent = 4
    oPan = 5
    oRecenter = 6
    oSingleSelect = 7
    oMultiSelect = 8
    oRectSelect = 9
    oPolySelect = 10
    oRadiusSelect = 11
    oPointBuffer = 27
    oInfo = 12
    oCreateLocationPoint = 13
    oCreateLocationLine = 14
    oCreateLocationMultipoint = 15
    oCreateLocationArea = 16
    oCreateLocationRadius = 17
    oCreateLocationPolyline = 18
    oCreateLocationText = 19
    oInfoDrillDown = 20
    oInfoOasisObject = 21
    oLineSelect = 22
    oPolyLineSelect = 23
    oAreaSelect = 24
    oFeatureSelect = 25
    oCircleSelect = 26
    oMeasure = 28
End Enum

Public Enum OASIS_GIS_DATA_SOURCE
    File = 0
    Database = 1
    Memory = 2
End Enum

Public Enum OASIS_GIS_DATA_TYPE
    vSHP = 0
    vTAB = 1
    vKML = 2
    vGPX = 3
    vNSQL = 4
    vOSQL = 5
    vGML = 6
    vDXF = 7
End Enum

Public Type ExportFile
    sFilter As String
    sDialogTitle As String
    sFileExtention As String
    fType As OASIS_GIS_DATA_TYPE
End Type

Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer ' e.g. = &h0031 = .31
    dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
    dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
    dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
    dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
    dwFileFlagsMask As Long ' = &h3F for version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
End Type
