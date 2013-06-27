VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmSpatialiseDD 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spatialise"
   ClientHeight    =   2925
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   7260
   Icon            =   "frmSpatialiseDD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7260
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   2925
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7260
      _cx             =   12806
      _cy             =   5159
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
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   2
      ChildSpacing    =   4
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
      GridRows        =   6
      GridCols        =   6
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmSpatialiseDD.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CommandButton cmdEnterMGRS 
         Caption         =   "Enter MGRS"
         Enabled         =   0   'False
         Height          =   270
         Left            =   5940
         TabIndex        =   19
         Top             =   180
         Width           =   1290
      End
      Begin VB.CommandButton cmdEnterXY 
         Caption         =   "Enter XY"
         Enabled         =   0   'False
         Height          =   270
         Left            =   4455
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   180
         Width           =   1425
      End
      Begin C1SizerLibCtl.C1Elastic C1ElasticZoom 
         Height          =   795
         Left            =   5940
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   810
         Width           =   1290
         _cx             =   2275
         _cy             =   1402
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
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
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
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.VScrollBar VScrollZoom 
            Height          =   495
            Left            =   480
            Max             =   2
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.Label lblZoom 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Zoom "
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Width           =   1095
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1ElasticPan 
         Height          =   810
         Left            =   5940
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1290
         _cx             =   2275
         _cy             =   1429
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
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
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
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.VScrollBar VScrollPan 
            Height          =   495
            Left            =   480
            Max             =   2
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Value           =   1
            Width           =   255
         End
         Begin VB.HScrollBar HScrollPan 
            Height          =   255
            Left            =   240
            Max             =   2
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   360
            Value           =   1
            Width           =   735
         End
         Begin VB.Label lblPan 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pan "
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdFlash 
         Caption         =   "Flash"
         Height          =   240
         Left            =   6615
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   510
         Width           =   615
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   240
         Left            =   5940
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   510
         Width           =   615
      End
      Begin VB.ListBox List2 
         Height          =   1620
         ItemData        =   "frmSpatialiseDD.frx":68F4
         Left            =   2985
         List            =   "frmSpatialiseDD.frx":68F6
         TabIndex        =   7
         Top             =   810
         Width           =   2895
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   360
         Left            =   2985
         TabIndex        =   6
         Top             =   2535
         Width           =   2895
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   360
         Left            =   30
         TabIndex        =   5
         Top             =   2535
         Width           =   2895
      End
      Begin VB.ListBox List1 
         Height          =   1620
         ItemData        =   "frmSpatialiseDD.frx":68F8
         Left            =   30
         List            =   "frmSpatialiseDD.frx":68FA
         TabIndex        =   4
         Top             =   810
         Width           =   2895
      End
      Begin VB.ComboBox cmbShapeType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSpatialiseDD.frx":68FC
         Left            =   1500
         List            =   "frmSpatialiseDD.frx":690C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   2895
      End
      Begin VB.Label lblY 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "y"
         Height          =   240
         Left            =   2985
         TabIndex        =   8
         Top             =   510
         Width           =   2895
      End
      Begin VB.Label lblPoints 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         Height          =   240
         Left            =   30
         TabIndex        =   3
         Top             =   510
         Width           =   2895
      End
      Begin VB.Label lblShapetype 
         BackStyle       =   0  'Transparent
         Caption         =   "  Shapetype"
         Height          =   270
         Left            =   30
         TabIndex        =   2
         Top             =   180
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmSpatialiseDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event CreateObject(sObjectType As String)
Public Event RestoreDDWindow(bCommitShape As Boolean)
Public Event FlashIt()
Public Event UpdateShape(oShape As TatukGIS_XDK10.XGIS_Shape)
Public Event ZoomIn(bZoomIn As Boolean)
Public Event PanLeftRight(bLeft As Boolean)
Public Event PanUpDown(bUp As Boolean)
Public Event ConvertMGRS(sMGRS As String, x As Double, y As Double)

Public bButtonPressed As Boolean
Private mLayer As TatukGIS_XDK10.XGIS_LayerVector
Private bDontUpdateList As Boolean

Public Function Init(bPointsOnly As Boolean)

cmbShapeType.Clear
cmbShapeType.AddItem "Point"
'cmbShapeType.Text = "Point"

If Not bPointsOnly Then

    cmbShapeType.AddItem "Polygon"
    cmbShapeType.AddItem "Polyline"
    cmbShapeType.AddItem "None"
 '   cmbShapeType.Text = "None"
    

End If

    

End Function

Public Function GetKMLForShape()

    Dim sRetval As String
    Dim i As Long
    
    Select Case cmbShapeType.Text
    
        Case "Point"
            sRetval = "<Point><coordinates>"
        Case "Polyline"
            sRetval = "<LineString><coordinates>"
        Case "Polygon"
            sRetval = "<Polygon><coordinates>"
        Case Else
            sRetval = ""
    
    End Select
    
    i = 0
    Do Until i = List1.ListCount
        sRetval = sRetval & List1.List(i) & ","
        sRetval = sRetval & List2.List(i) & " "
        i = i + 1
    Loop
        
    Select Case cmbShapeType.Text
    
        Case "Point"
            sRetval = sRetval & "<coordinates><Point>"
        Case "Polyline"
            sRetval = sRetval & "<coordinates><LineString>"
        Case "Polygon"
            sRetval = sRetval & "<coordinates><Polygon>"
        Case Else
            sRetval = ""
    
    End Select
    
    GetKMLForShape = sRetval

End Function

Public Sub SetLayer(oLayer As TatukGIS_XDK10.XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo SetLayer_Err
        '</EhHeader>
    
        Dim oShape As TatukGIS_XDK10.XGIS_Shape
        Dim sPoints As String
    
100     Set mLayer = oLayer
102     Set oShape = oLayer.GetShape(oLayer.GetLastUid)
104     cmbShapeType.Enabled = False
106     bDontUpdateList = False
    
108     If oShape Is Nothing Then
110         MsgBox "No spatial component found - there may be corruption in your database", vbInformation
            Unload Me
        Else

116         Do Until i = oShape.GetNumPoints

118             If Len(sPoints) > 1 Then
120                 sPoints = sPoints & ";" & oShape.GetPoint(0, i).x & ";" & oShape.GetPoint(0, i).y
                Else
122                 sPoints = oShape.GetPoint(0, i).x & ";" & oShape.GetPoint(0, i).y
                End If

124             i = i + 1

            Loop

126         SetListValues sPoints
128         DetectComboItem
130         RaiseEvent UpdateShape(oShape)

        End If
    
        '<EhFooter>
        Exit Sub

SetLayer_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.SetLayer " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetListValues(sValues As String)
        '<EhHeader>
        On Error GoTo SetListValues_Err
        '</EhHeader>

        Dim sValueArray() As String
        Dim i As Integer

100     sValueArray = Split(sValues, ";")
102     List1.Clear
104     List2.Clear
106     i = 0

108     Do Until i >= UBound(sValueArray)
    
110         List1.AddItem sValueArray(i)
112         List2.AddItem sValueArray(i + 1)
114         i = i + 2
        Loop
    
        cmdEnterXY.Enabled = False
        cmdEnterMGRS.Enabled = False
116     cmbShapeType.Enabled = False
118     C1ElasticPan.Enabled = True
120     C1ElasticZoom.Enabled = True
122     cmdFlash.Enabled = True
124     C1ElasticPan.BackColor = vbWhite
126     C1ElasticZoom.BackColor = vbWhite
        cmbShapeType.BackColor = C1Elastic1.BackColor

        '<EhFooter>
        Exit Sub

SetListValues_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.SetListValues " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub DetectComboItem()
        '<EhHeader>
        On Error GoTo DetectComboItem_Err
        '</EhHeader>
    
100     bDontUpdateList = True
    
102     If List1.ListCount = 1 Then
104         cmbShapeType.Text = "Point"
106     ElseIf List1.ListCount = 0 Then
108         cmbShapeType.Text = "None"
110     ElseIf List1.List(0) = List1.List(List1.ListCount - 1) And List2.List(0) = List2.List(List2.ListCount - 1) Then
112         cmbShapeType.Text = "Polygon"
        Else
114         cmbShapeType.Text = "Polyline"
        End If
    
116     bDontUpdateList = False
    
        '<EhFooter>
        Exit Sub

DetectComboItem_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.DetectComboItem " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmbShapeType_Click()
        '<EhHeader>
        On Error GoTo cmbShapeType_Click_Err
        '</EhHeader>
           
        Dim oShape As TatukGIS_XDK10.XGIS_Shape
           
100     If Not bDontUpdateList Then
    
102         List1.Clear
104         List2.Clear

106         RaiseEvent CreateObject(cmbShapeType.Text)
           
108         Select Case cmbShapeType.Text

                Case "Point"
                    cmdEnterXY.Enabled = True
                    cmdEnterMGRS.Enabled = True
110                 Set oShape = mLayer.CreateShape(XgisShapeTypePoint)

112             Case "Polyline"
                    cmdEnterXY.Enabled = False
                    cmdEnterMGRS.Enabled = False
114                 Set oShape = mLayer.CreateShape(XgisShapeTypeArc)

116             Case "Polygon"
                    cmdEnterXY.Enabled = False
                    cmdEnterMGRS.Enabled = False
118                 Set oShape = mLayer.CreateShape(XgisShapeTypePolygon)
            End Select

120         If Not mLayer.FileInfo = "Generic Vector Layer" Then mLayer.SaveData
        
        End If

        '<EhFooter>
        Exit Sub

cmbShapeType_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.cmbShapeType_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     bButtonPressed = False
102     Unload Me
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdEdit_Click()
        '<EhHeader>
        On Error GoTo cmdEdit_Click_Err
        '</EhHeader>
        
        If cmbShapeType.Text = "Point" Then
            cmdEnterXY.Enabled = True
            cmdEnterMGRS.Enabled = True
        Else
            cmdEnterXY.Enabled = False
            cmdEnterMGRS.Enabled = False
        End If
        
100     cmbShapeType.Enabled = True
102     C1ElasticPan.Enabled = False
104     C1ElasticZoom.Enabled = False
106     cmdFlash.Enabled = False
108     C1ElasticPan.BackColor = Me.BackColor
110     C1ElasticZoom.BackColor = Me.BackColor
        cmbShapeType.BackColor = vbWhite
112     Call cmbShapeType_Click
    
        '<EhFooter>
        Exit Sub

cmdEdit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.cmdEdit_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdEnterMGRS_Click()

    Dim x As Double
    Dim y As Double
    Dim sMGRS As String
    Dim sRetval As String

    sRetval = InputBox("Please enter MGRS coordinate", "Get MGRS", "")

    If Len(sRetval) > 1 Then

        sMGRS = sRetval
        RaiseEvent ConvertMGRS(sMGRS, x, y)
        
        Dim oShape As New TatukGIS_XDK10.XGIS_Shape
        Dim oPt As New TatukGIS_XDK10.XGIS_Point

        mLayer.GetShape(mLayer.GetLastUid).Delete

        Set oShape = mLayer.CreateShape(XgisShapeTypePoint)
        oShape.AddPart

        oPt.Prepare x, y
        oShape.AddPoint oPt
        If Not mLayer.FileInfo = "Generic Vector Layer" Then mLayer.SaveData
        SetLayer mLayer
 
    End If
    
End Sub

Private Sub cmdEnterXY_Click()

    Dim x As Double
    Dim y As Double
    Dim sRetval As String

    sRetval = InputBox("Please enter X coordinate", "Get X", "0")

    If IsNumeric(sRetval) Then

        x = CDbl(sRetval)
        sRetval = InputBox("Please enter Y coordinate", "Get Y", "0")
        
        If IsNumeric(sRetval) Then

            y = CDbl(sRetval)
            
            Dim oShape As New TatukGIS_XDK10.XGIS_Shape
            Dim oPt As New TatukGIS_XDK10.XGIS_Point

            mLayer.GetShape(mLayer.GetLastUid).Delete

            Set oShape = mLayer.CreateShape(XgisShapeTypePoint)
            oShape.AddPart

            oPt.Prepare x, y
            oShape.AddPoint oPt
            If Not mLayer.FileInfo = "Generic Vector Layer" Then mLayer.SaveData
            SetLayer mLayer
 
        End If
 
    End If

End Sub

Private Sub cmdFlash_Click()
        '<EhHeader>
        On Error GoTo cmdFlash_Click_Err
        '</EhHeader>
100     RaiseEvent FlashIt
        '<EhFooter>
        Exit Sub

cmdFlash_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.cmdFlash_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOk_Click()
        '<EhHeader>
        On Error GoTo cmdOk_Click_Err
        '</EhHeader>

100     If List1.ListCount = 0 Then
102         MsgBox "You have not set the spatial information for this record!", vbCritical
        Else
104         bButtonPressed = True
106         Me.Hide
108         RaiseEvent RestoreDDWindow(True)
        End If

        '<EhFooter>
        Exit Sub

cmdOk_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.cmdOK_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     bButtonPressed = False
        cmbShapeType.BackColor = C1Elastic1.BackColor
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     RaiseEvent RestoreDDWindow(False)
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub List1_Click()
        '<EhHeader>
        On Error GoTo List1_Click_Err
        '</EhHeader>
100     List2.ListIndex = List1.ListIndex
        '<EhFooter>
        Exit Sub

List1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.List1_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub List2_Click()
        '<EhHeader>
        On Error GoTo List2_Click_Err
        '</EhHeader>
100     List1.ListIndex = List2.ListIndex
        '<EhFooter>
        Exit Sub

List2_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.List2_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function GetSpatialDataString() As String
        '<EhHeader>
        On Error GoTo GetSpatialDataString_Err
        '</EhHeader>

        Dim i As Long
        Dim sRetval As String
100     i = 0
    
102     Do Until i = List1.ListCount
    
104         If Len(sRetval) > 1 Then
106             sRetval = sRetval & ";" & List1.List(i) & "," & List2.List(i)
            Else
108             sRetval = List1.List(i) & "," & List2.List(i)
            End If

110         i = i + 1
        Loop
    
112     GetSpatialDataString = sRetval

        '<EhFooter>
        Exit Function

GetSpatialDataString_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.GetSpatialDataString " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub HScrollPan_Change()
        '<EhHeader>
        On Error GoTo HScrollPan_Change_Err
        '</EhHeader>

100     With HScrollPan
    
102         If .Value = 0 Then
                ' MsgBox "left"
104             RaiseEvent PanLeftRight(True)
106         ElseIf .Value = 2 Then
                ' MsgBox "right"
108             RaiseEvent PanLeftRight(False)
            End If
    
110         .Value = 1
        
        End With

        '<EhFooter>
        Exit Sub

HScrollPan_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.HScrollPan_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub VScrollPan_Change()
        '<EhHeader>
        On Error GoTo VScrollPan_Change_Err
        '</EhHeader>

100     With VScrollPan
    
102         If .Value = 0 Then
                'MsgBox "up"
104             RaiseEvent PanUpDown(True)
106         ElseIf .Value = 2 Then
                'MsgBox "down"
108             RaiseEvent PanUpDown(False)
            End If
    
110         .Value = 1
    
        End With
    
        '<EhFooter>
        Exit Sub

VScrollPan_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.VScrollPan_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub VScrollZoom_Change()
        '<EhHeader>
        On Error GoTo VScrollZoom_Change_Err
        '</EhHeader>

100     With VScrollZoom
    
102         If .Value = 0 Then
                'MsgBox "in"
104             RaiseEvent ZoomIn(True)
            
106         ElseIf .Value = 2 Then
                'MsgBox "out"
108             RaiseEvent ZoomIn(False)
            End If
    
110         .Value = 1
    
        End With
    
        '<EhFooter>
        Exit Sub

VScrollZoom_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialiseDD.VScrollZoom_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
