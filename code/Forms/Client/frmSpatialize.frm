VERSION 5.00
Begin VB.Form frmSpatialize 
   Caption         =   "Spatialiser"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin OASISClient.OASISLocator OASISLocator1 
      Height          =   3465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   6112
   End
End
Attribute VB_Name = "frmSpatialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event CheckCoordInMap(x As Double, y As Double, sMGRS As String)
Public Event PcodePicked()
Public Event FocusRecieved(Name As String)
Public Event LocRadiusSelectTool()
Public Event GetCoordFromMapTool()
Public Event hglAdminLevel0()
Public Event hglAdminLevel1()
Public Event hglAdminLevel2()
Public Event hglAdminLevel3()
Public Event hglAdminLocation()
Public Event LocationFound(sName As String)
Public Event RestoreDDWindow()
Dim mRS As adodb.Recordset
Dim bDynamicData As Boolean

Public Sub SpatialiseDD(PassedGISViewer As TatukGIS_XDK9.XGIS_Viewer, _
                        oRS As adodb.Recordset)
    
    bDynamicData = True
    
    Dim bSuccess As Boolean
    Dim x As Integer
    Dim y As Double
    Dim xVal As Double
    Dim yVal As Double
    bSuccess = False
    Set mRS = oRS
    ' Set mFRM = oFRM

    OASISLocator1.Init PassedGISViewer

    With mRS.Fields("XMIN")

        If Not IsNull(.Value) And Not .Value = "" And Not .Value = 0 Then

            xVal = .Value
            bSuccess = True
        End If

    End With

    With mRS.Fields("YMIN")

        If Not IsNull(.Value) And Not .Value = "" And Not .Value = 0 Then

            yVal = .Value
            bSuccess = True
        Else
            bSuccess = False
        End If

    End With

    If bSuccess Then OASISLocator1.SetXY CStr(xVal), CStr(yVal), False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If bDynamicData Then

        mRS.Fields("XMIN").Value = Me.OASISLocator1.Longitude
        mRS.Fields("XMAX").Value = Me.OASISLocator1.Longitude
        mRS.Fields("YMIN").Value = Me.OASISLocator1.Latitude
        mRS.Fields("YMAX").Value = Me.OASISLocator1.Latitude
        mRS.Fields("SHAPETYPE").Value = 2
        RaiseEvent RestoreDDWindow
    
    End If

    bDynamicData = False
    
End Sub

Private Sub OASISLocator1_CheckCoordInMap(x As Double, _
                                          y As Double, _
                                          sMGRS As String)
    RaiseEvent CheckCoordInMap(x, y, sMGRS)
End Sub

Private Sub OASISLocator1_FocusRecieved(Name As String)
    RaiseEvent FocusRecieved(Name)
End Sub

Private Sub OASISLocator1_GetCoordFromMapTool()
    RaiseEvent GetCoordFromMapTool
End Sub

Private Sub OASISLocator1_LocationFound(sName As String)
    RaiseEvent LocationFound(sName)
End Sub

Private Sub OASISLocator1_LocRadiusSelectTool()
    RaiseEvent LocRadiusSelectTool
End Sub

Private Sub OASISLocator1_PcodePicked()
    RaiseEvent PcodePicked
End Sub
