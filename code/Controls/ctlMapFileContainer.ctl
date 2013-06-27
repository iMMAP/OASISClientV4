VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F8F9FBF9-12B5-11D4-8ED3-00E07D815373}#1.0#0"; "MBScroll.ocx"
Begin VB.UserControl ctlMapFileContainer 
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   ScaleHeight     =   6180
   ScaleWidth      =   3825
   ToolboxBitmap   =   "ctlMapFileContainer.ctx":0000
   Begin VB.PictureBox picExport 
      Height          =   735
      Left            =   240
      ScaleHeight     =   675
      ScaleWidth      =   2295
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.PictureBox pic 
      Height          =   495
      Left            =   1140
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   1380
      Visible         =   0   'False
      Width           =   1215
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6180
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   3825
      _cx             =   6747
      _cy             =   10901
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
      GridRows        =   2
      GridCols        =   3
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"ctlMapFileContainer.ctx":0312
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   300
         Left            =   0
         Picture         =   "ctlMapFileContainer.ctx":0367
         TabIndex        =   5
         Top             =   5880
         Width           =   3825
      End
      Begin MBScroller.Scroller Scroller1 
         Height          =   5880
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   10372
         ScrollBars      =   2
         Begin OASISClient.ctlMapFilePointer ctlMapFilePointer1 
            Height          =   1155
            Index           =   0
            Left            =   180
            TabIndex        =   3
            Top             =   240
            Visible         =   0   'False
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   2037
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   1560
      Picture         =   "ctlMapFileContainer.ctx":1CA29
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   2160
      Width           =   615
   End
End
Attribute VB_Name = "ctlMapFileContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event MapLoad(sPath As String, oExtent As XGIS_Extent)
Public Event MapSave(sPath As String, sGUID As String)
Public Event MapDelete(sGUID As String)
Public Event MapNew()
Public Event MapInfo(sInfo As String, sGUID As String)
Public Event MapPreview(oPicture As StdPicture)

Private lTopOfNextControl As Long
Private mCnn As ADODB.Connection
Private mRS As ADODB.Recordset

Public IsMapSelected As Boolean
Private m_CurrentMapIndex As Long
Private m_CurrentMapPath As String
Private m_LoadedMapIndex As Long
Private m_LoadedMapPath As String

Private m_ScrollerScrolling As Boolean
Private m_CurrentMapGUID As String

Public Sub HighlightMapAsActive(lIndex As Long, _
                                sGUID As String, _
                                Optional sPath As String = "")
    
    If lIndex = -1 Then lIndex = m_CurrentMapIndex
    m_LoadedMapIndex = lIndex
    ctlMapFilePointer1(lIndex).SetTitle Replace(ctlMapFilePointer1(lIndex).GetTitle, " (active)", "")
    ctlMapFilePointer1(lIndex).SetActive
    ctlMapFilePointer1(lIndex).SetTitle ctlMapFilePointer1(lIndex).GetTitle & " (active)"
    
    m_CurrentMapGUID = sGUID
    m_CurrentMapPath = sPath

    If sPath = "" Then
        m_LoadedMapPath = ctlMapFilePointer1(lIndex).GetFilePath
    Else
        m_LoadedMapPath = Replace$(sPath, "CLIENTDBPATH", g_sAppPath)
    End If
   
    Scroller1.Refresh
    IsMapSelected = True

End Sub

Private Sub C1Elastic1_RealignFrame()
        '<EhHeader>
        On Error GoTo C1Elastic1_RealignFrame_Err
        '</EhHeader>
        Dim i As Integer
    
100     Do Until i = ctlMapFilePointer1.Count
102         ctlMapFilePointer1(i).Width = Scroller1.Width - 315
104         i = i + 1
        Loop
       
        '<EhFooter>
        Exit Sub

C1Elastic1_RealignFrame_Err:
        Err.Raise vbObjectError + 100, "OASISClient.ctlMapFileContainer.C1Elastic1_RealignFrame", "ctlMapFileContainer component failure"
        '</EhFooter>
End Sub

Private Sub cmdDelete_Click()
    RaiseEvent MapDelete(m_CurrentMapGUID)
End Sub

Private Sub cmdEdit_Click()
    RaiseEvent MapSave(m_CurrentMapPath, m_CurrentMapGUID)
End Sub

Private Sub cmdNew_Click()
     RaiseEvent MapNew
     ctlMapFilePointer1(m_LoadedMapIndex).LoadMap
End Sub

Public Sub SetActiveMapAsUserMap()
    ctlMapFilePointer1(m_LoadedMapIndex).SetMapType UserCustom
End Sub

Private Sub ctlMapFilePointer1_MapDelete(Index As Integer, _
                                         sGUID As String)

    If Not ctlMapFilePointer1.Count = 2 Then
    
        m_CurrentMapIndex = Index
        If m_LoadedMapIndex = Index Then m_LoadedMapIndex = 0
        RaiseEvent MapDelete(sGUID)
        
    Else
        MsgBox "Sorry, you cannot remove the last map in the map library!", vbInformation
    End If

End Sub

Private Sub ctlMapFilePointer1_MapEdit(Index As Integer, sGUID As String)
    m_CurrentMapIndex = Index
    RaiseEvent MapSave(m_CurrentMapPath, sGUID)
End Sub

Private Sub ctlMapFilePointer1_MapInfo(Index As Integer, _
                                       sMapInfo As String, sGUID As String)
        '<EhHeader>
        On Error GoTo ctlMapFilePointer1_MapInfo_Err
        '</EhHeader>
        m_CurrentMapIndex = Index
100     RaiseEvent MapInfo(sMapInfo, sGUID)
        '<EhFooter>
        Exit Sub

ctlMapFilePointer1_MapInfo_Err:
        Err.Raise vbObjectError + 100, "OASISClient.ctlMapFileContainer.ctlMapFilePointer1_MapInfo", "ctlMapFileContainer component failure"
        '</EhFooter>
End Sub

Public Sub PrepareMapForLoad(ByRef sMapPath As String)

        Dim oStream As ADODB.Stream
        Dim oRS As New ADODB.Recordset
        Dim sText As String
112     Set oRS = New ADODB.Recordset
114     oRS.Open "SELECT * FROM [ttkGISProjectDef] where sGUID = '" & Replace(sMapPath, "Database Driven:", "") & "'", mCnn, adOpenDynamic, adLockBatchOptimistic

116     If Not oRS.EOF Then
                   
118         m_CurrentMapGUID = oRS.Fields("sGUID").value
            ' sText = IIf(Len(oRS.Fields("MapData").Value) = 0, "", Replace(oRS.Fields("MapData").Value, "CLIENTDBPATH", g_sAppPath))

120         If Not IsNull(oRS.Fields("MapData").value) Then
122             sText = Replace(oRS.Fields("MapData").value, "CLIENTDBPATH", g_sAppPath)
            Else
124             MsgBox "There was a problem loading the map project from the database", vbCritical
                Exit Sub
            End If

126         sMapPath = g_sAppPath & "\data\user\maps\" & oRS.Fields("sGUID").value & ".ttkgp"
128         m_CurrentMapPath = Replace(sMapPath, "Database Driven:", "")
                
130         If Not IsNull(oRS.Fields("MapData").value) Then
132             ctlMapFilePointer1(m_CurrentMapIndex).SetInfo oRS.Fields("sInfo").value
            End If
                
134         oRS.Close
                
            On Error Resume Next
136         Kill sMapPath
            ' On Error GoTo LoadProjectFileFromDB_Err
                
138         Set oStream = New ADODB.Stream
140         oStream.Open
142         oStream.Type = 2    ' Set type to text
144         oStream.Charset = "ascii"
146         oStream.WriteText sText
148         oStream.SaveToFile (sMapPath)
150         oStream.Close
152         Set oStream = Nothing


                
        End If
            
164     Set oRS = Nothing
End Sub

Private Sub ctlMapFilePointer1_MapLoad(Index As Integer, _
                                       ByVal sMapPath As String, _
                                       bIsFileBased As Boolean)
        '<EhHeader>
        On Error GoTo ctlMapFilePointer1_MapLoad_Err
        '</EhHeader>

        'On Error GoTo LoadProjectFileFromDB_Err

        If bIsFileBased Then
            If Not FileExists(sMapPath) Then
                MsgBox "The file specified does not exist.  Please contact an OASIS administrator", vbInformation
                Exit Sub
            End If
        End If
        
100     If IsMapSelected Then ctlMapFilePointer1(m_LoadedMapIndex).SetInactive
104     If IsMapSelected Then ctlMapFilePointer1(m_LoadedMapIndex).SetTitle Replace(ctlMapFilePointer1(m_LoadedMapIndex).GetTitle, " (active)", "")
106     ctlMapFilePointer1(Index).SetTitle ctlMapFilePointer1(Index).GetTitle & " (active)"
102     ctlMapFilePointer1(Index).SetActive

108     m_CurrentMapIndex = Index
        m_LoadedMapIndex = Index
        IsMapSelected = True
        
110     If Not bIsFileBased And Not Len(Replace(sMapPath, "Database Driven:", "")) = Len(sMapPath) Then
            PrepareMapForLoad sMapPath
        End If

        If Not bIsFileBased Then
166         RaiseEvent MapLoad(sMapPath, ctlMapFilePointer1(Index).GetExtent)
        Else
            RaiseEvent MapLoad(sMapPath, Nothing)
        End If
        
        g_bDefaultMapChanged = True
        g_bDefaultMapChangedGUID = ctlMapFilePointer1(Index).MapGUID

        '<EhFooter>
        Exit Sub

ctlMapFilePointer1_MapLoad_Err:
 MsgBox Err.Description & vbCrLf & " in OASISClient.ctlMapFileContainer.ctlMapFilePointer1_MapLoad (" & Erl & ") " & Err.Description

        'Err.Raise vbObjectError + 100, "OASISClient.ctlMapFileContainer.ctlMapFilePointer1_MapLoad", "ctlMapFileContainer component failure"
        '</EhFooter>
End Sub

Public Sub SavePreviewToBlob()
    'TODO this bugger... It is painful....ARE YOU HERE?
    pic.Picture = Clipboard.GetData

    If pic.Picture <> 0 Then
        Clipboard.Clear
        pic.Picture = LoadPicture("")
    End If
    
End Sub

Private Sub ctlMapFilePointer1_MapPreview(Index As Integer, _
                                          oPicture As StdPicture)
        '<EhHeader>
        On Error GoTo ctlMapFilePointer1_MapPreview_Err

        '</EhHeader>
        m_CurrentMapIndex = Index
        
100     If Not oPicture Is Nothing Then
102         RaiseEvent MapPreview(oPicture)
        Else
104         MsgBox "There is no map preview!!!"
        End If

        '<EhFooter>
        Exit Sub

ctlMapFilePointer1_MapPreview_Err:
        Err.Raise vbObjectError + 100, "OASISClient.ctlMapFileContainer.ctlMapFilePointer1_MapPreview", "ctlMapFileContainer component failure"
        '</EhFooter>
End Sub

Public Sub SetCurrentExtent(oExtent As XGIS_Extent)
    ctlMapFilePointer1(m_CurrentMapIndex).SetExtent oExtent
End Sub

Public Sub Init()
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        Dim lHeightOfControl As Long
        Dim lWidthOfControl As Long
        Dim sName As String
        Dim sPath As String
        Dim lBackColour As Long
        Dim lForeColour As Long
        Dim bUGMap As Boolean
        Dim bIsFileBased As Boolean
        Dim oExtent As XGIS_Extent
        Dim iMapIndex As Long
        Dim MapType As MapTypeValue
        
100     Scroller1.vValue = 0

102     Do Until ctlMapFilePointer1.Count = 1
104         Unload ctlMapFilePointer1(ctlMapFilePointer1.Count - 1)
        Loop
    
106     If g_sAppPath = "" Then Exit Sub

108     lTopOfNextControl = 1
110     lHeightOfControl = 1665
112     lWidthOfControl = 3901 '3301
114     lWidthOfControl = Scroller1.Width '* 0.9

116     UserControl.Width = lWidthOfControl
    
118     Set mCnn = m_Cnn
120     Set mRS = New ADODB.Recordset
122     mRS.Open "SELECT * FROM [ttkGISProjectDef] order by [InUse], [sName]", mCnn, adOpenDynamic, adLockReadOnly

124     With mRS
            
126         iMapIndex = 1

128         Do Until mRS.EOF
            
130             If .Fields("InUse").value Then
132                 lBackColour = RGB(255, 255, 255)
134                 lForeColour = vbRed
136                 sName = HandleNullString(.Fields("sName").value) & " (active)"
                Else
138                 lBackColour = Scroller1.BackColor
140                 lForeColour = vbBlack
142                 sName = HandleNullString(.Fields("sName").value)
                End If
            
144             If .Fields("bSavedToDB").value = True Then
146                 sPath = "Database Driven:" & HandleNullString(.Fields("sGUID").value)
                Else
148                 sPath = Replace$(HandleNullString(.Fields("sFilePath").value), "CLIENTDBPATH", g_sAppPath)
                End If
                
150             If Not IsNull(.Fields("oImagePreview").value) Then
152                 picExport = GetPictureFromRecordset(mRS, "oImagePreview")
154                 picExport.Width = picExport.Picture.Width
156                 picExport.Height = picExport.Picture.Height
                Else
158                 Set picExport = cmdNew.Picture
160                 picExport.Width = picExport.Picture.Width
162              picExport.Height = picExport.Picture.Height
                    
                End If
                
164             bUGMap = .Fields("bUGMap").value
166             bIsFileBased = Not .Fields("bSavedToDb").value
                
168             If bIsFileBased Then
170                 MapType = FileBased
                Else

172                 If bUGMap Then
174                     MapType = UserGroup
                    Else
176                     MapType = UserCustom
                    End If
                End If
                
178             Set oExtent = New XGIS_Extent
180             oExtent.Prepare .Fields("XMIN").value, .Fields("YMIN").value, .Fields("XMAX").value, .Fields("YMAX").value
            
182             AddControl .Fields("sGUID").value, lWidthOfControl, lHeightOfControl, HandleNullString(sName), HandleNullString(sPath), HandleNullString(.Fields("sInfo").value), lBackColour, lForeColour, MapType, oExtent, picExport

184             If .Fields("InUse").value Then
186                 HighlightMapAsActive m_CurrentMapIndex, .Fields("sGUID").value
                End If

188             mRS.MoveNext
            
            Loop
        
        End With
        
190     If 1 = 1 Or lTopOfNextControl > Scroller1.Height Then
192         m_ScrollerScrolling = True
194         Call C1Elastic1_RealignFrame
        End If
        
196     Scroller1.Refresh
    
198     mRS.Close
        '<EhFooter>
        Exit Sub

Init_Err:
        'needed to remove error handling here - for some reason it was tirggering error on oasis close
        ''MsgBox "Error in OASISClient.ctlMapFileContainer.Init (" & Erl & ") " & Err.Description
        '</EhFooter>
End Sub

Public Function GetActiveMapName() As String
    GetActiveMapName = ctlMapFilePointer1(m_LoadedMapIndex).GetTitle
End Function

Public Function GetActiveMapPath() As String
    GetActiveMapPath = m_LoadedMapPath  ' ctlMapFilePointer1(m_LoadedMapIndex).GetFilePath
End Function

Public Sub SetCurrentActiveMapAsInactive()

    If IsMapSelected Then
        ctlMapFilePointer1(m_LoadedMapIndex).SetInactive
        m_LoadedMapIndex = -1
        IsMapSelected = False
    End If

End Sub


Public Function GetActiveMapExtent() As XGIS_Extent
'On Error Resume Next
Set GetActiveMapExtent = ctlMapFilePointer1(m_LoadedMapIndex).GetExtent
End Function

Public Property Get picDC() As Long
    picDC = picExport.hdc
End Property

Public Function picPicture() As PictureBox
    Set picPicture = picExport
End Function

Public Sub SetExportImageSize(pic As PictureBox)

    With picExport
        If Not pic Is Nothing Then
            .Move .left, .top, pic.Width, pic.Height
        End If
        
        .ZOrder 0
        .Cls
        .AutoRedraw = True
    End With

End Sub

Public Sub SaveImageToDB(ppic As Picture, RS As ADODB.Recordset, _
                         pColName As String)
    
    Dim pb As PropertyBag
    Set pb = New PropertyBag
    pb.WriteProperty "MyImage", ppic
    RS.Fields(pColName).AppendChunk pb.Contents
    RS.UpdateBatch adAffectCurrent
    Set pb = Nothing
    
End Sub

Private Function GetPictureFromRecordset(RS As ADODB.Recordset, _
                                         pColName As String) As Picture
    Dim pb As PropertyBag
    Set pb = New PropertyBag
    pb.Contents = RS.Fields(pColName).GetChunk(RS.Fields(pColName).ActualSize)
    Set GetPictureFromRecordset = pb.ReadProperty("MyImage")
    
    Set pb = Nothing
    
End Function

Public Sub SaveActiveMap(oImage As Picture, sInfo As String, sName As String, sGUID As String)
   
    Set mRS = New ADODB.Recordset
    
    If sGUID = "" Then
        mRS.Open "SELECT * FROM [ttkGISProjectDef] where sGUID = '" & m_CurrentMapGUID & "'", mCnn, adOpenDynamic, adLockBatchOptimistic
    Else
        mRS.Open "SELECT * FROM [ttkGISProjectDef] where sGUID = '" & sGUID & "'", mCnn, adOpenDynamic, adLockBatchOptimistic
    End If
    
    With mRS
    
        If Not mRS.EOF Then
            SaveImageToDB oImage, mRS, "oImagePreview"
            ctlMapFilePointer1(m_CurrentMapIndex).SetImage oImage     'picExport.Picture  '.Picture         'oImage
            ctlMapFilePointer1(m_CurrentMapIndex).SetInfo sInfo
            ctlMapFilePointer1(m_CurrentMapIndex).SetTitle sName
        End If
        
    End With
    
    mRS.Close
    Set mRS = Nothing

End Sub

Public Sub AddControl(sGUID As String, _
                      lWidth As Long, _
                      lHeight As Long, _
                      sName As String, _
                      sPath As String, _
                      sInfo As String, _
                      oBackColor As ColorConstants, _
                      oCaptionColour As ColorConstants, _
                      MapTypePassed As MapTypeValue, _
                      oExtent As XGIS_Extent, _
                      Optional oImage As StdPicture)
                      
                      
        '<EhHeader>
        On Error GoTo AddControl_Err
        '</EhHeader>

        Dim ctlDynamic As Control
        Dim lNumControls As Long
        Dim bDisableRefresh As Boolean
    
        If Not lWidth = -1 Then bDisableRefresh = True
        If lWidth = -1 Then lWidth = Scroller1.Width - 315
        If lHeight = -1 Then lHeight = 1665
        
100     lNumControls = lNumControls + 1
102     lNumControls = ctlMapFilePointer1.Count
104     Load ctlMapFilePointer1(lNumControls)
106     Set ctlDynamic = ctlMapFilePointer1(lNumControls - 1)

        If lNumControls > 1 Then
            ctlDynamic.Move 1, ctlMapFilePointer1(lNumControls - 2).Height + ctlMapFilePointer1(lNumControls - 2).top, lWidth, lHeight
        Else
108         ctlDynamic.Move 1, lTopOfNextControl, lWidth, lHeight
        End If
        
        ctlDynamic.SetMapType MapTypePassed
        
       ' ctlDynamic.SetUGMap bUGMap
       ' IsFileBased = True

       ' If Left$(sPath, 16) = "Database Driven:" Then
      '      ctlDynamic.SetFileBased False
      '  Else
      '      ctlDynamic.SetFileBased True
      '  End If

110     ctlDynamic.Visible = True
112     lTopOfNextControl = lTopOfNextControl + lHeight
114     ctlDynamic.SetTitle sName
116     ctlDynamic.SetFilePath sPath

118     If Not oImage Is Nothing Then ctlDynamic.SetImage oImage
120     ctlDynamic.SetBackColour oBackColor
122     ctlDynamic.SetCaptionColour oCaptionColour
124     ctlDynamic.SetInfo sInfo
        
        ctlDynamic.SetExtent oExtent
        ctlDynamic.oGUID sGUID
        m_CurrentMapIndex = ctlDynamic.Index
    
        If Not bDisableRefresh Then

            If lTopOfNextControl > Scroller1.Height Then
                m_ScrollerScrolling = True
                Call C1Elastic1_RealignFrame
                Scroller1.Refresh
            End If
        
        End If
 
        '<EhFooter>
        Exit Sub

AddControl_Err:
        Err.Raise vbObjectError + 100, "OASISClient.ctlMapFileContainer.AddControl", "ctlMapFileContainer component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_Initialize()
        '<EhHeader>
        On Error GoTo UserControl_Initialize_Err
        '</EhHeader>
100     Init

        'HaversineDistance 100, 0, 100, 0
        '<EhFooter>
        Exit Sub

UserControl_Initialize_Err:
        'Err.Raise vbObjectError + 100, _
         "OASISClient.ctlMapFileContainer.UserControl_Initialize", _
         "ctlMapFileContainer component failure"
        '</EhFooter>
End Sub
