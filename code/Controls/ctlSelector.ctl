VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlSelector 
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12240
   ScaleHeight     =   3105
   ScaleWidth      =   12240
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12240
      _cx             =   21590
      _cy             =   5477
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   0
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "      Please wait - your data is loading..."
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
      GridRows        =   8
      GridCols        =   8
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"ctlSelector.ctx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.ListBox lstLayerName 
         Height          =   255
         Left            =   1050
         TabIndex        =   21
         Top             =   2280
         Width           =   2910
      End
      Begin VB.Frame Frame1 
         Caption         =   "Selection Method"
         Height          =   2685
         Left            =   3990
         TabIndex        =   15
         Top             =   0
         Width           =   3345
         Begin VB.ComboBox cmbOtherLayer 
            Enabled         =   0   'False
            Height          =   315
            Left            =   600
            TabIndex        =   22
            Text            =   "Combo1"
            Top             =   1980
            Width           =   2355
         End
         Begin VB.OptionButton OptAnotherLayer 
            Caption         =   "Another layer"
            Height          =   195
            Left            =   300
            TabIndex        =   20
            Top             =   1680
            Width           =   1275
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Polygon"
            Height          =   195
            Left            =   300
            TabIndex        =   19
            Top             =   1350
            Width           =   1815
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Polyline"
            Height          =   195
            Left            =   300
            TabIndex        =   18
            Top             =   1020
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Circle (specify km)"
            Height          =   195
            Left            =   300
            TabIndex        =   17
            Top             =   690
            Width           =   1755
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Circle (draw)"
            Height          =   195
            Left            =   300
            TabIndex        =   16
            Top             =   360
            Width           =   1755
         End
      End
      Begin VB.ListBox lstLayer 
         Height          =   2205
         Left            =   0
         TabIndex        =   14
         Top             =   375
         Width           =   3960
      End
      Begin XpressEditorsLibCtl.dxPickEdit COMLayer 
         Height          =   315
         Left            =   1050
         OleObjectBlob   =   "ctlSelector.ctx":00DD
         TabIndex        =   13
         Top             =   0
         Width           =   2910
      End
      Begin VB.CommandButton cmdGO 
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "Segoe UI Symbol"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11595
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Clear Selections"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   645
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   390
         Left            =   7365
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1005
         Visible         =   0   'False
         Width           =   990
         _cx             =   1746
         _cy             =   688
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
         BackColor       =   12648447
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
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
         GridCols        =   2
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   0
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"ctlSelector.ctx":017D
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic lblOtherLayerName 
            Height          =   390
            Left            =   0
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   510
            _cx             =   900
            _cy             =   688
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
            BackColor       =   12648447
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "layername"
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
         End
      End
      Begin XpressEditorsLibCtl.dxTextEdit txtBufferDistance 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   8385
         OleObjectBlob   =   "ctlSelector.ctx":01B6
         TabIndex        =   3
         Top             =   0
         Width           =   945
      End
      Begin VB.CheckBox chkDistance 
         Caption         =   "   Calculate distance (km)"
         Height          =   345
         Left            =   9360
         TabIndex        =   2
         Top             =   0
         Value           =   1  'Checked
         Width           =   2205
      End
      Begin XpressEditorsLibCtl.dxPickEdit ComSelectBy 
         Height          =   315
         Left            =   9360
         OleObjectBlob   =   "ctlSelector.ctx":0226
         TabIndex        =   1
         Top             =   2715
         Width           =   2205
      End
      Begin C1SizerLibCtl.C1Elastic C1EBufferKm 
         Height          =   345
         Left            =   7365
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   990
         _cx             =   1746
         _cy             =   609
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
         Caption         =   "  Buffer (km)"
         Align           =   0
         AutoSizeChildren=   0
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
      End
      Begin C1SizerLibCtl.C1Elastic C1ESelectBy 
         Height          =   390
         Left            =   8385
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2715
         Width           =   945
         _cx             =   1667
         _cy             =   688
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
         Caption         =   "  Select by"
         Align           =   0
         AutoSizeChildren=   0
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
      End
      Begin C1SizerLibCtl.C1Elastic C1EActiveLayer 
         Height          =   345
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   1020
         _cx             =   1799
         _cy             =   609
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
         Caption         =   "  Active layer"
         Align           =   0
         AutoSizeChildren=   0
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
      End
      Begin MSComctlLib.ListView lstFields 
         Height          =   2070
         Left            =   7365
         TabIndex        =   7
         Top             =   615
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   3651
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin C1SizerLibCtl.C1Elastic C1EFieldsUsed 
         Height          =   210
         Left            =   7365
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   375
         Width           =   4875
         _cx             =   8599
         _cy             =   370
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
         BackColor       =   12648447
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "Fields used"
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
      End
      Begin CONTROLSLibCtl.dxProgressBar dxProgressBar1 
         Height          =   390
         Left            =   0
         TabIndex        =   11
         Top             =   2715
         Width           =   8355
         _Version        =   65536
         _cx             =   14737
         _cy             =   688
         ForeColor       =   0
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MinPos          =   0
         MaxPos          =   100
         Pos             =   50
         Step            =   10
         ShowText        =   -1  'True
         Orientation     =   0
         StartColor      =   16711680
         EndColor        =   16777215
         DrawBorderStyle =   1
         ShowTextStyle   =   0
         DrawBarStyle    =   2
         DrawBarBorderStyle=   2
      End
   End
End
Attribute VB_Name = "ctlSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event ChangeSelector(sTools As String)
Public Event DisableSelections(sLayerName As String)
Public Event FlashShape(sUID As String)

Public Event GetLayers(sLayers As String)
Public Event UseOtherLayer(sOtherLayer As String)
Public Event MergeOtherSelections(sOtherLayer As String, bForceAll As Boolean, bSuccess As Boolean)
Public Event ChangeActiveLayer(sName As String, sExcluded As String)
Public Event GetLayerSelectedInLegend(sLayerName As String)

Private m_PErcentMark As Long
Private m_PErcentJump As Long
Private m_LayerName As String

Public Sub AddNewLayer(sName As String, _
                       sCaption As String)
    'COMLayer.items.Add sCaption
    lstLayer.AddItem sCaption
End Sub

Public Function GetActiveLayer() As String
    GetActiveLayer = m_LayerName
End Function

Public Sub RenewSelection()

    If Not ComSelectBy = "-- Select tool --" Then
        ComSelectBy = "-- Select tool --"
    End If

End Sub

Public Function DistanceEnabled() As Boolean
    DistanceEnabled = IIf(chkDistance.value = vbChecked, True, False)
End Function

Public Function ToggleGridVisible(bVisible As Boolean)
    dxProgressBar1.Pos = 0

End Function

Public Function GetFields() As String
        '<EhHeader>
        On Error GoTo GetFields_Err
        '</EhHeader>
    
        Dim i As Long
100     i = 0
    
102     Do Until i = lstFields.ListItems.Count

104         If lstFields.ListItems(i + 1).Checked Then
106             GetFields = GetFields & ";" & lstFields.ListItems(i + 1).Text
            End If

108         i = i + 1
        Loop
    
        '<EhFooter>
        Exit Function

GetFields_Err:
        MsgBox Err.Description & vbCrLf & "in ctlSelector.GetFields_Err " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function GetTool() As OASIS_TOOLS

    Select Case ComSelectBy

        Case "Circle (specify km)"
        
            GetTool = oPointBuffer

        Case "Polyline"
            
            GetTool = oLineSelect
        
        Case "Circle (draw)"
            
            GetTool = oCircleSelect

        Case "Polygon"
           
            GetTool = oAreaSelect
           
        Case Else
            
            GetTool = 0
            
    End Select
    
End Function

Public Sub InitProgressBar(lCount As Long)
        '<EhHeader>
        On Error GoTo InitProgressBar_Err
        '</EhHeader>
100     dxProgressBar1.Step = 1
102     dxProgressBar1.MaxPos = lCount + 1
104     dxProgressBar1.MinPos = 0
106     dxProgressBar1.Pos = 0
108     m_PErcentMark = Round((lCount + 1) / 100, 0)
110     m_PErcentJump = m_PErcentMark
        '<EhFooter>
        Exit Sub

InitProgressBar_Err:
        MsgBox Err.Description & vbCrLf & "in ctlSelector.InitProgressBar_Err " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub ProgressStep()
    dxProgressBar1.DoStep

    If dxProgressBar1.Pos = m_PErcentMark Then

        DoEvents
        m_PErcentMark = m_PErcentMark + m_PErcentJump
    End If

End Sub

Public Sub SetLayer(sLayerName As String, _
                    sCaption As String, _
                    sFields As String)
        '<EhHeader>
        On Error GoTo SetLayer_Err
        '</EhHeader>
    
100     If Len(sLayerName) > 0 Then
            Dim sFieldNames() As String
            Dim i As Long
102         i = 1
    
104         If Len(m_LayerName) > 0 And Not ComSelectBy = "Layer with existing selection(s)" And Not ComSelectBy = "Layer with all its feature(s)" Then RaiseEvent DisableSelections(m_LayerName)
106         ComSelectBy.Enabled = True
    
108         'ComLayer = sCaption
110         m_LayerName = sLayerName
112         sFieldNames = Split(sFields, ",")
114         lstFields.ListItems.Clear
        
116         Do Until i > UBound(sFieldNames)
118             lstFields.ListItems.Add , , sFieldNames(i)

120             If i < 3 And sFieldNames(i) <> "GUID1" Then lstFields.ListItems(i).Checked = True
122             i = i + 1
            Loop

124         lstFields.Enabled = True
    
        Else
126         'ComLayer = ""
128         RenewSelection
130         ComSelectBy.Enabled = False
132         lstFields.ListItems.Clear
134         lstFields.Enabled = False
        
        End If

        '<EhFooter>
        Exit Sub

SetLayer_Err:
        MsgBox Err.Description & vbCrLf & "in ctlSelector.SetLayer_Err " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdGO_Click()

    If Trim(lblOtherLayerName) <> "" Then
    
        If lblOtherLayerName <> m_LayerName Then
            RaiseEvent UseOtherLayer(lblOtherLayerName)
            RenewSelection
        Else
            MsgBox "You need to select a new layer in the drop-down box to analyse"
        End If
    
    Else
        MsgBox "You have no layer selected in the drop-down box!"
    End If

End Sub

Private Sub ComSelectBy_Change()
        '<EhHeader>
        On Error GoTo ComSelectBy_Change_Err
        '</EhHeader>

        Dim sLayers As String
        Dim i As Long
        Dim sLayersArray() As String
        Dim sLayerName As String
        Dim bSuccess As Boolean
    
100     If Not ComSelectBy = "Layer with existing selection(s)" Then
102         RaiseEvent ChangeSelector(ComSelectBy)
        Else
104         RaiseEvent ChangeSelector("Pan")
        End If

106     If ComSelectBy = "Polyline" Or ComSelectBy = "Circle (specify km)" Then
108         txtBufferDistance.Enabled = True
        Else
110         txtBufferDistance.Enabled = False
        End If
 
112     If ComSelectBy = "Layer with existing selection(s)" Then
114         cmdGO.Visible = True
116         RaiseEvent MergeOtherSelections(m_LayerName, False, bSuccess)

            If bSuccess = False Then cmdGO.Visible = False
118     ElseIf ComSelectBy = "Layer with all its feature(s)" Then
            cmdGO.Visible = True
            RaiseEvent MergeOtherSelections(GetActiveLayer, True, bSuccess)

120         If bSuccess = False Then cmdGO.Visible = False

        Else
122         cmdGO.Visible = False
        End If

        '<EhFooter>
        Exit Sub

ComSelectBy_Change_Err:
        Err.Raise vbObjectError + 100, "OASISClient.ctlSelector.ComSelectBy_Change", "ctlSelector component failure"
        '</EhFooter>
End Sub

Public Function GetBuffer() As Double
        '<EhHeader>
        On Error GoTo GetBuffer_Err
        '</EhHeader>

100     GetBuffer = 0

102     If Not txtBufferDistance.Enabled Then
104         GetBuffer = -1
106     ElseIf Len(txtBufferDistance) > 0 Then
108         GetBuffer = CDbl(txtBufferDistance) / 100

        End If

        '<EhFooter>
        Exit Function

GetBuffer_Err:
        MsgBox Err.Description & vbCrLf & "in ctlSelector.GetBuffer_Err " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function GetLayerListBox() As ListBox
    Set GetLayerListBox = lstLayer 'COMLayer
End Function

Private Sub lstLayer_Click()
        Dim sLayerName As String
        Dim sExcluded As String

100     SafeMoveFirst g_RSGISGridTableSettings
102     g_RSGISGridTableSettings.Find "alias = '" & lstLayer.Text & "'"
    
104     If Not g_RSGISGridTableSettings.EOF And Not g_RSGISGridTableSettings.Bof Then
106         sLayerName = g_RSGISGridTableSettings.Fields("Name").value
108         sExcluded = IIf(IsNull(g_RSGISGridTableSettings.Fields("excludedFlds").value), "", g_RSGISGridTableSettings.Fields("excludedFlds").value)
        Else
110         sLayerName = lstLayer.Text 'COMLayer
        End If

112     'RaiseEvent DisableSelections(m_LayerName)

114     If lstLayer.Text = "---Nothing---" Then
            RaiseEvent DisableSelections(m_LayerName)
            m_LayerName = sLayerName
116         RenewSelection
            
118         ComSelectBy.Enabled = False
120         lstFields.ListItems.Clear
122         lstFields.Enabled = False
        Else

124         If ComSelectBy = "Layer with existing selection(s)" Or ComSelectBy = "Layer with all its feature(s)" Then
126             RaiseEvent ChangeSelector("Pan")
            Else
128             RaiseEvent DisableSelections(m_LayerName)
            End If
            
        End If

        m_LayerName = sLayerName
      
130     RaiseEvent ChangeActiveLayer(sLayerName, sExcluded)
End Sub

Private Sub UserControl_Initialize()
        '<EhHeader>
        On Error GoTo UserControl_Initialize_Err
        '</EhHeader>

100     ComSelectBy.items.Add "-- Select tool --"
102     ComSelectBy.items.Add "Circle (draw)"
104     ComSelectBy.items.Add "Circle (specify km)"
106     ComSelectBy.items.Add "Polygon"
108     ComSelectBy.items.Add "Polyline"
        ComSelectBy.items.Add "Layer with all its feature(s)"
110     ComSelectBy.items.Add "Layer with existing selection(s)"
112     txtBufferDistance = 0
114     RenewSelection
    
        '<EhFooter>
        Exit Sub

UserControl_Initialize_Err:
        MsgBox Err.Description & vbCrLf & "in ctlSelector.UserControl_Initialize_Err " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

