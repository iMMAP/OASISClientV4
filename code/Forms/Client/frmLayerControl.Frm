VERSION 5.00
Begin VB.Form LayerControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Layer Control"
   ClientHeight    =   3345
   ClientLeft      =   1470
   ClientTop       =   3615
   ClientWidth     =   6450
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmLayerControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3345
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton btnClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame frmReorder 
      Caption         =   "Reorder"
      Height          =   1215
      Left            =   3360
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
      Begin VB.CommandButton cmdDown 
         Caption         =   "&Down"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "&Up"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frmLayers 
      Caption         =   "Layers"
      Height          =   1215
      Left            =   1800
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add..."
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frmProperties 
      Caption         =   "Properties"
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
      Begin VB.CommandButton cmdLabel 
         Caption         =   "&Label..."
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Display..."
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox ckAutoLabel 
      Caption         =   "A&utomatic Labels"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CheckBox ckSelectable 
      Caption         =   "&Selectable"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox ckVisible 
      Caption         =   "&Visible"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.ListBox lstLayers 
      Height          =   1425
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblLayers 
      Caption         =   "La&yers:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "LayerControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim iPrevSel As Integer
'Private gGIS As Map
'
'Public Sub Init(oGIS As TatukGIS_XDK9.XGIS_Viewer)
'  Set gGIS = oGIS
'  InitLayers
'  FormToCenter LayerControl.hwnd
'  LayerControl.Show 1
'End Sub
'
'Private Sub AddDeletedLayer(ByVal ind As Integer)
'  ReDim Preserve DeletedLayers(UBound(DeletedLayers) + 1)
'  DeletedLayers(UBound(DeletedLayers)) = gMap.Layers(ind).Name
'End Sub
'
'Private Sub FillLayers(ld() As LayerDescr, Map As Map)
'  Dim LN, i As Integer
'
'  LN = Map.Layers.Count
'  ReDim ld(LN)
'
'  For i = 1 To LN
'    ld(i).Visible = Map.Layers(i).Visible
'    ld(i).Selectable = Map.Layers(i).Selectable
'    ld(i).AutoLabel = False
'    ld(i).InitPos = i
'    ld(i).LabelChanged = False
'    ld(i).StyleChanged = False
'    ld(i).AutoLabel = Map.Layers(i).AutoLabel
'    ld(i).ZoomLayer = Map.Layers(i).ZoomLayer
'    ld(i).ZoomMin = Map.Layers(i).ZoomMin
'    ld(i).ZoomMax = Map.Layers(i).ZoomMax
'    ld(i).Name = Map.Layers(i).Name
'    ld(i).Override = Map.Layers(i).OverrideStyle
'    Set ld(i).StyleProp = Map.Layers(i).Style
'    Set ld(i).LabelProp = Map.Layers(i).LabelProperties
'  Next
'  ReDim DeletedLayers(0)
'End Sub
'
'Private Sub MoveLayerUp(ld() As LayerDescr, lstBox As ListBox, ByVal Lay As Integer)
'  Dim sItem As String, t As LayerDescr
'
'  t = ld(Lay)
'  ld(Lay) = ld(Lay - 1)
'  ld(Lay - 1) = t
'
'  sItem = lstBox.List(Lay - 1)
'  lstBox.RemoveItem Lay - 1
'  lstBox.AddItem sItem, Lay - 2
'  lstBox.ListIndex = Lay - 2
'End Sub
'
'Private Sub MoveLayerDown(ld() As LayerDescr, lstBox As ListBox, ByVal Lay As Integer)
'  Dim sItem As String, t As LayerDescr
'
'  t = ld(Lay)
'  ld(Lay) = ld(Lay + 1)
'  ld(Lay + 1) = t
'
'  sItem = lstBox.List(Lay - 1)
'  lstBox.RemoveItem Lay - 1
'  lstBox.AddItem sItem, Lay
'  lstBox.ListIndex = Lay
'End Sub
'
'Private Sub RemoveLayer(ld() As LayerDescr, lstBox As ListBox, ByVal Lay As Integer)
'  Dim i As Integer
'
'  lstBox.RemoveItem Lay - 1
'  For i = Lay To UBound(ld) - 1
'    ld(i) = ld(i + 1)
'  Next
'  ReDim Preserve ld(UBound(ld) - 1)
'End Sub
'
'Private Sub ShowLayerSettings(ld() As LayerDescr, ByVal Lay As Integer, ckVis As CheckBox, ckSel As CheckBox, ckLab As CheckBox)
'  If ld(Lay).Visible Then
'    ckVis.Value = 1
'  Else
'    ckVis.Value = 0
'  End If
'  If ld(Lay).Selectable Then
'    ckSel.Value = 1
'  Else
'    ckSel.Value = 0
'  End If
'  If ld(Lay).AutoLabel Then
'    ckLab.Value = 1
'  Else
'    ckLab.Value = 0
'  End If
'End Sub
'
'Private Sub InitLayers()
'  Dim i As Integer
'
'  FillLayers ld, gMap
'  lstLayers.Clear
'  For i = 1 To gMap.Layers.Count
'    lstLayers.AddItem (gMap.Layers(i).Name)
'  Next
'  lstLayers.Selected(0) = True
'  iPrevSel = 0
'  ShowLayerSettings ld, 1, ckVisible, ckSelectable, ckAutoLabel
'  cmdUp.Enabled = False
'End Sub
'
'Private Sub btnClose_Click()
'  LayerControl.Hide
'End Sub
'
'Private Sub btnOK_Click()
'  Dim i As Integer
'
'  Hide
'
'' Turn off the screen updating
'  gMap.AutoRedraw = False
'
'' To remove all the deleted layers
'  For i = 1 To UBound(DeletedLayers)
''    gMap.Layers.Remove left$(DeletedLayers(i), Len(DeletedLayers(i)))
'    gMap.Layers.Remove DeletedLayers(i)
'  Next
'
'' To add all new layers
'  For i = 1 To UBound(ld)
'    If ld(i).NewLayer Then
'      ld(i).Name = gMap.Layers.Add(ld(i).Name).Name
'    End If
'  Next
'
'' To reorder the layers
'  For i = 1 To UBound(ld)
'    If gMap.Layers(i).Name <> ld(i).Name Then
'      gMap.Layers.Move GetLayerPosition(ld(i).Name), i
'    End If
'  Next
'
'' To set the styles of the layers
'  For i = 1 To UBound(ld)
'    gMap.Layers(i).Visible = ld(i).Visible
'    gMap.Layers(i).Selectable = ld(i).Selectable
'    gMap.Layers(i).AutoLabel = ld(i).AutoLabel
'    gMap.Layers(i).OverrideStyle = ld(i).Override
'    gMap.Layers(i).ZoomLayer = ld(i).ZoomLayer
'    gMap.Layers(i).ZoomMin = ld(i).ZoomMin
'    gMap.Layers(i).ZoomMax = ld(i).ZoomMax
'  Next
'
'' Redraw the control
'  gMap.AutoRedraw = True
'End Sub
'
'Private Sub ckAutoLabel_Click()
'  ld(lstLayers.ListIndex + 1).AutoLabel = (ckAutoLabel.Value = 1)
'End Sub
'
'Private Sub ckSelectable_Click()
'  ld(lstLayers.ListIndex + 1).Selectable = (ckSelectable.Value = 1)
'End Sub
'
'Private Sub ckVisible_Click()
'  ld(lstLayers.ListIndex + 1).Visible = (ckVisible.Value = 1)
'End Sub
'
'Private Sub cmdAdd_Click()
'  Dim ind As Integer, i As Integer, bCancel As Boolean
'
'  dlgSpecifyLayer.DialogTitle = "Specify layer table"
'  dlgSpecifyLayer.Filter = "MapInfo Tables (*.tab)|*.tab|All Files(*.*)|*.*"
'  dlgSpecifyLayer.FilterIndex = 1
'  bCancel = False
'  On Error GoTo CancelErr
'  dlgSpecifyLayer.ShowOpen
'
'  If bCancel Then
'    Exit Sub
'  End If
'
'  ind = lstLayers.ListIndex + 1
'  ReDim Preserve ld(UBound(ld) + 1)
'  For i = UBound(ld) - 1 To ind Step -1
'    ld(i + 1) = ld(i)
'  Next
'  ld(ind).Name = dlgSpecifyLayer.FileName
'  ld(ind).NewLayer = True
'  ld(ind).Visible = True
'  ld(ind).Selectable = True
'  ld(ind).AutoLabel = False
'  ld(i).ZoomLayer = False
'  ld(i).ZoomMin = 0
'  ld(i).ZoomMax = 0
'  ld(ind).InitPos = 0
'  Set ld(ind).LabelProp = gMap.Layers(1).LabelProperties
'  Set ld(ind).StyleProp = gMap.DefaultStyle
'  ld(ind).LabelChanged = True
'  lstLayers.AddItem ld(ind).Name, ind - 1
'  lstLayers.ListIndex = ind - 1
'  Exit Sub
'CancelErr:
'  bCancel = True
'  Resume Next
'End Sub
'
'Private Sub cmdDisplay_Click()
'  DisplayProperties.Activate lstLayers.ListIndex + 1
'End Sub
'
'Private Sub cmdDown_Click()
'  Dim Lay As Integer
'
'  Lay = lstLayers.ListIndex + 1
'  If Lay = lstLayers.ListCount Then
'    Beep
'    Exit Sub
'  End If
'  MoveLayerDown ld, lstLayers, Lay
'  If Lay = lstLayers.ListCount - 1 Then
'    cmdDown.Enabled = False
'  End If
'End Sub
'
'Private Sub cmdLabel_Click()
'  Dim ind As Integer
'
'  ind = lstLayers.ListIndex + 1
'  LabelProps.Activate gMap, ind
''  ld(ind).Name , ld(ind).NewLayer, ld(ind).LabelProp.Style, ld(ind).LabelProp.Position, ld(ind).LabelProp.Field, ld(ind).LabelProp.Parallel
'End Sub
'
'Private Sub cmdRemove_Click()
'  Dim Lay As Integer
'
'  Lay = lstLayers.ListIndex + 1
'
'  If Not ld(Lay).NewLayer Then
'    AddDeletedLayer ld(Lay).InitPos
'  End If
'
'  RemoveLayer ld, lstLayers, Lay
'
'  If lstLayers.ListCount = Lay - 1 Then
'    lstLayers.ListIndex = Lay - 2
'    cmdDown.Enabled = False
'  Else
'    lstLayers.ListIndex = Lay - 1
'  End If
'  ShowLayerSettings ld, lstLayers.ListIndex + 1, ckVisible, ckSelectable, ckAutoLabel
'End Sub
'
'Private Sub cmdUp_Click()
'  Dim Lay As Integer
'
'  Lay = lstLayers.ListIndex + 1
'  If Lay = 1 Then
'    Beep
'    Exit Sub
'  End If
'  MoveLayerUp ld, lstLayers, Lay
'  If Lay = 2 Then
'    cmdUp.Enabled = False
'  End If
'End Sub
'
'Private Sub Form_Load()
'
'End Sub
'
'Private Sub lstLayers_Click()
'  If lstLayers.SelCount > 0 Then
'    ShowLayerSettings ld, lstLayers.ListIndex + 1, ckVisible, ckSelectable, ckAutoLabel
'  End If
'  If lstLayers.ListCount = lstLayers.ListIndex + 1 Then
'    cmdDown.Enabled = False
'  Else
'    cmdDown.Enabled = True
'  End If
'  If lstLayers.ListIndex = 0 Then
'    cmdUp.Enabled = False
'  Else
'    cmdUp.Enabled = True
'  End If
'End Sub
'
'Private Sub lstLayers_DblClick()
'  cmdDisplay_Click
'End Sub
'
'Private Function GetLayerPosition(ByVal LN As String) As Integer
'  Dim i As Integer
'
'  For i = 1 To gMap.Layers.Count
'    If gMap.Layers(i).Name = LN Then
'      GetLayerPosition = i
'      Exit Function
'    End If
'  Next
'  GetLayerPosition = 0
'End Function
Private Sub Form_Load()
If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
End Sub
