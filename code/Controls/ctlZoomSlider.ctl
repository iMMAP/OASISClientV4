VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.UserControl ctlZoomSlider 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000D&
   BackStyle       =   0  'Transparent
   ClientHeight    =   10290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   ScaleHeight     =   10290
   ScaleWidth      =   3630
   Begin C1SizerLibCtl.C1Elastic C1ZoomIn 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Width           =   375
      _cx             =   661
      _cy             =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
      Caption         =   "+"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   -1  'True
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   0
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin C1SizerLibCtl.C1Elastic ZoomBarBottom 
      Height          =   3735
      Left            =   960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   960
      _cx             =   1693
      _cy             =   6588
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      FloodColor      =   16711680
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   4
      FloodPercent    =   25
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   0
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin C1SizerLibCtl.C1Elastic ZoomBarTop 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   960
      _cx             =   1693
      _cy             =   6588
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      FloodColor      =   16711680
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   8
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   3
      FloodPercent    =   25
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   400
      MinChildSize    =   1
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   9
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   1
      FrameWidth      =   0
      FrameColor      =   -2147483627
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"ctlZoomSlider.ctx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.Shape ZoomTick 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   45
      Left            =   1920
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      X1              =   360
      X2              =   600
      Y1              =   8760
      Y2              =   9120
   End
End
Attribute VB_Name = "ctlZoomSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event ZoomChanged(oExtent As XGIS_Extent)
Private CurrentZoom As Integer
Private ZoomTicks() As Shape
Private bLinesDrawn As Boolean
Private m_Extents() As XGIS_Extent
Private m_MaxInitExtent As XGIS_Extent
Private lZoomSelectColour As Long
Private lZoomBackColour As Long
Private lZoomCrosshairColour As Long
Private bOperationInProgress As Boolean

Public Sub SetColours(oZoomSelectColour As Long, _
                      oZoomBackColour As Long, _
                      oZoomCrosshairColour As Long)
    lZoomSelectColour = oZoomSelectColour
    lZoomBackColour = oZoomBackColour
    lZoomCrosshairColour = oZoomCrosshairColour
    UpdateBars
End Sub


Public Sub Init(o_MaxInitExtent As XGIS_Extent, _
                o_CurrentInitExtent As XGIS_Extent, _
                oZoomSelectColour As Long, _
                oZoomBackColour As Long, _
                oZoomCrosshairColour As Long)
                
    'NEEDED
    Dim i As Integer
    Dim iInterval As Integer
    Dim ctlLine As Line
    ReDim Preserve ZoomTicks(20)
    lZoomSelectColour = oZoomSelectColour
    lZoomBackColour = oZoomBackColour
    lZoomCrosshairColour = oZoomCrosshairColour

    If Not bLinesDrawn Then

        Do Until i > 19
            Set ZoomTicks(i) = Controls.Add("VB.Shape", "ShapeControl" & i + 1, Me)
            ZoomTicks(i).ZOrder 0
            i = i + 1
        Loop

    End If
        
    bLinesDrawn = True
    UpdateBars
                
    ReDim m_Extents(19)
    Set m_MaxInitExtent = o_MaxInitExtent
    SetZoomLevels o_MaxInitExtent
    UserControl.BackColor = vbWhite
    SetZoomPointerFromExtent o_CurrentInitExtent ', False
    
End Sub

Private Sub SetZoomLevels(oReferenceExtent As XGIS_Extent)
        '<EhHeader>
        On Error GoTo SetZoomLevels_Err
        '</EhHeader>
        
        'NEEDED
        Dim iZoomValue As Integer
        Dim dDistanceToXCentroid As Double
        Dim dDistanceToYCentroid As Double
        Dim dCentroidX As Double
        Dim dCentroidY As Double
        Dim dNextStepX  As Double
        Dim dNextStepY  As Double
      
100     dDistanceToXCentroid = Abs((oReferenceExtent.xmax + 360000) - (oReferenceExtent.xmin + 360000)) / 2
102     dDistanceToYCentroid = Abs((oReferenceExtent.ymax + 360000) - (oReferenceExtent.ymin + 360000)) / 2
104     dCentroidX = oReferenceExtent.xmax - dDistanceToXCentroid
106     dCentroidY = oReferenceExtent.ymax - dDistanceToYCentroid
   
108     dStepx = dDistanceToXCentroid / 19
110     dStepy = dDistanceToYCentroid / 19
    
112     For iZoomValue = 1 To 19

            If iZoomValue = 19 Then
                dNextStepX = Abs(LogBase(18, 19) - 1) * dDistanceToXCentroid / 2
                dNextStepY = Abs(LogBase(18, 19) - 1) * dDistanceToYCentroid / 2
            Else

114             dNextStepX = Abs(LogBase(iZoomValue, 19) - 1) * dDistanceToXCentroid
116             dNextStepY = Abs(LogBase(iZoomValue, 19) - 1) * dDistanceToYCentroid

            End If

118         Set m_Extents(20 - iZoomValue) = New XGIS_Extent
120         m_Extents(20 - iZoomValue).Prepare dCentroidX - (dNextStepX), dCentroidY - (dNextStepY), dCentroidX + (dNextStepX), dCentroidY + (dNextStepY)
            Debug.Print iZoomValue & " -- " & dNextStepX
        Next

        '<EhFooter>
        Exit Sub

SetZoomLevels_Err:
        Err.Raise vbObjectError + 100, "OASISClient.ctlZoomSlider.SetZoomLevels", "ctlZoomSlider component failure"
        '</EhFooter>
End Sub

Private Function LogBase(iNum As Integer, iBase As Integer) As Double
    LogBase = Log(iNum) / Log(iBase)
    If LogBase = 0 Then LogBase = 0.01
End Function

Private Sub UpdateBars()

    'NEEDED
    Dim i As Integer
    Dim iInterval As Integer
     
    If bLinesDrawn Then
        ZoomTicks(0).Visible = True
        ZoomTicks(0).FillStyle = 0
        ZoomTicks(0).left = (UserControl.Width / 2) - 75
        ZoomTicks(0).top = 0
        ZoomTicks(0).Width = 150
        ZoomTicks(0).Height = UserControl.Height
        ZoomTicks(0).FillColor = lZoomBackColour
        
        ZoomTicks(0).BorderColor = lZoomCrosshairColour
        i = 1
        iInterval = UserControl.Height / 20
        bBold = True

        Do Until i = 20
            ZoomTicks(i).Visible = True
            ZoomTicks(i).top = 1 'Lines(i).BorderWidth * 2
            ZoomTicks(i).left = 0
            ZoomTicks(i).Width = UserControl.Width
            ZoomTicks(i).FillStyle = 0
            ZoomTicks(i).top = i * iInterval
            ZoomTicks(i).Height = 45
            ZoomTicks(i).FillColor = lZoomCrosshairColour
        ZoomTicks(i).BorderColor = lZoomCrosshairColour
            i = i + 1
        Loop
        
        ZoomTicks(0).ZOrder 0
        
    End If
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    'NEEDED
    bOperationInProgress = True
    Dim lPercentage As Integer
    lPercentage = Round((100 * Y / UserControl.Height), 0)
    CurrentZoom = GetNearestZoomValue(lPercentage)

    If bLinesDrawn Then SetZoomPointerFromCurrentZoomLevel
    RaiseEvent ZoomChanged(m_Extents(CurrentZoom))
    bOperationInProgress = False

End Sub

Private Sub UserControl_Resize()
    'NEEDED
    UpdateBars
End Sub

Private Function GetNearestZoomValue(iPercentage As Integer) As Integer
    
    'NEEDED
    Dim i As Integer
    Dim iCurrentValue As Integer
    Dim iNearestValue As Integer
    
    i = 0
    iCurrentValue = 1
    iNearestValue = 1000
    GetNearestZoomValue = 10

    Do Until i = 20
    
        If Abs(iPercentage - iNearestValue) > Abs(iPercentage - iCurrentValue) Then
            iNearestValue = iCurrentValue
            GetNearestZoomValue = i
        End If
        
        i = i + 1
        iCurrentValue = (100 / 20) * i
    
    Loop
    
    If GetNearestZoomValue < 1 Then GetNearestZoomValue = 1
    If GetNearestZoomValue > 19 Then GetNearestZoomValue = 19
    
End Function

Private Function GetNearestZoomValueFromExtent(ByVal oExtent As XGIS_Extent) As Integer
    
    'NEEDED
    Dim i As Integer
    Dim oExtentTemp As New XGIS_Extent
    Dim oExtentTemp2 As New XGIS_Extent

    i = 1
    oExtentTemp.Prepare oExtent.xmin, oExtent.ymin, oExtent.xmin + Abs(m_Extents(i).xmax - m_Extents(i).xmin), oExtent.ymin + Abs(m_Extents(i).ymax - m_Extents(i).ymin)
        
    Do Until Not GisUtils.GisIsContainExtent(oExtentTemp, oExtent)

        GetNearestZoomValueFromExtent = i
        i = i + 1

        If i > 19 Then Exit Do
        oExtentTemp.Prepare oExtent.xmin, oExtent.ymin, oExtent.xmin + Abs(m_Extents(i).xmax - m_Extents(i).xmin), oExtent.ymin + Abs(m_Extents(i).ymax - m_Extents(i).ymin)
       
    Loop
    
    If i < 20 Then

        oExtentTemp2.Prepare oExtent.xmin, oExtent.ymin, oExtent.xmin + Abs(m_Extents(i).xmax - m_Extents(i).xmin), oExtent.ymin + Abs(m_Extents(i).ymax - m_Extents(i).ymin)
        
        If Abs(oExtent.xmax - oExtentTemp2.xmax) < Abs(oExtent.xmax - oExtentTemp.xmax) Then
            GetNearestZoomValueFromExtent = i
        End If
        
    End If
    
    Set oExtentTemp = Nothing
    Set oExtentTemp2 = Nothing
    
End Function

Public Sub SetZoomPointerFromExtent(oExtent As XGIS_Extent)

    'NEEDED
    If Not bOperationInProgress Then
        CurrentZoom = GetNearestZoomValueFromExtent(oExtent)
        SetZoomPointerFromCurrentZoomLevel
    End If
    
End Sub

Private Sub SetZoomPointerFromCurrentZoomLevel()

    'NEEDED
    Shape1.FillColor = vbWhite
    Shape1.BorderColor = vbBlack
    Shape1.BorderWidth = 1
    Shape1.Height = (ZoomTicks(4).top - ZoomTicks(3).top) * 0.75
    Shape1.top = UserControl.Height - ((20 - CurrentZoom) * (UserControl.Height / 20)) - (Shape1.Height) / 2
    Shape1.Width = UserControl.Width
    Shape1.left = 0
    Shape1.ZOrder 0
    Shape1.FillColor = lZoomSelectColour
    Shape1.Visible = True
    Shape1.BorderColor = lZoomSelectColour
    Shape1.Refresh
    
    Line1.X1 = (UserControl.Width * 0.2)
    Line1.X2 = UserControl.Width - (UserControl.Width * 0.2)
    Line1.Y1 = UserControl.Height - ((20 - CurrentZoom) * (UserControl.Height / 20)) - 5
    Line1.Y2 = UserControl.Height - ((20 - CurrentZoom) * (UserControl.Height / 20)) - 5
    Line1.BorderColor = vbWhite
    Line1.ZOrder 0
    Line1.Refresh
    
End Sub

Public Function GetZoom() As Integer
    'NEEDED
    GetZoom = CurrentZoom
End Function

Public Sub ZoomIn()

    'NEEDED
    bOperationInProgress = True

    If CurrentZoom > 1 Then
        CurrentZoom = CurrentZoom - 1

        If bLinesDrawn Then SetZoomPointerFromCurrentZoomLevel
        RaiseEvent ZoomChanged(m_Extents(CurrentZoom))
    End If

    bOperationInProgress = False
End Sub

Public Sub ZoomOut()

    'NEEDED
    bOperationInProgress = True

    If CurrentZoom < 19 Then
        CurrentZoom = CurrentZoom + 1

        If bLinesDrawn Then SetZoomPointerFromCurrentZoomLevel
        RaiseEvent ZoomChanged(m_Extents(CurrentZoom))
    End If

    bOperationInProgress = False

End Sub

Private Sub UserControl_Terminate()

    'NEEDED
    Erase ZoomTicks
    Erase m_Extents
    Set m_MaxInitExtent = Nothing

End Sub


