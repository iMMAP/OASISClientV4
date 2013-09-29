VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmLocator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OASIS Locator"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2490
   Icon            =   "frmLocator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLocator.frx":6852
   ScaleHeight     =   4125
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1ElasticBAck 
      Height          =   4125
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2490
      _cx             =   4392
      _cy             =   7276
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
      BackColor       =   -2147483633
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
      GridRows        =   2
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmLocator.frx":D0A4
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CheckBox chkAutozoom 
         Caption         =   "Autozoom"
         Height          =   330
         Left            =   30
         TabIndex        =   14
         Top             =   3765
         Value           =   1  'Checked
         Width           =   2430
      End
      Begin C1SizerLibCtl.C1Elastic elAdmin 
         Height          =   3675
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   2430
         _cx             =   4286
         _cy             =   6482
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
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   6
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
         GridRows        =   13
         GridCols        =   3
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmLocator.frx":D0E4
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.ComboBox ComAdmin5 
            Enabled         =   0   'False
            Height          =   315
            Left            =   90
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2625
            Width           =   2250
         End
         Begin VB.ComboBox ComAdmin3 
            Enabled         =   0   'False
            Height          =   315
            Left            =   90
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1470
            Width           =   2250
         End
         Begin VB.ComboBox ComAdmin4 
            Enabled         =   0   'False
            Height          =   315
            Left            =   90
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2055
            Width           =   2250
         End
         Begin VB.ComboBox ComAdmin1 
            Height          =   315
            Left            =   90
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   315
            Width           =   2250
         End
         Begin VB.ComboBox ComAdmin6 
            Enabled         =   0   'False
            Height          =   315
            Left            =   90
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   3210
            Width           =   2250
         End
         Begin VB.ComboBox ComAdmin2 
            Enabled         =   0   'False
            Height          =   315
            Left            =   90
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   885
            Width           =   2250
         End
         Begin VB.Label lblAdmin5 
            AutoSize        =   -1  'True
            Caption         =   "District:"
            Height          =   210
            Left            =   90
            TabIndex        =   13
            Top             =   2415
            Width           =   2250
         End
         Begin VB.Label lblAdmin4 
            AutoSize        =   -1  'True
            Caption         =   "District:"
            Height          =   225
            Left            =   90
            TabIndex        =   12
            Top             =   1830
            Width           =   2250
         End
         Begin VB.Label lblAdmin3 
            AutoSize        =   -1  'True
            Caption         =   "Province:"
            Height          =   225
            Left            =   90
            TabIndex        =   11
            Top             =   1245
            Width           =   2250
         End
         Begin VB.Label lblAdmin1 
            AutoSize        =   -1  'True
            Caption         =   "Province:"
            Height          =   225
            Left            =   90
            TabIndex        =   10
            Top             =   90
            Width           =   2250
         End
         Begin VB.Label lblAdmin6 
            AutoSize        =   -1  'True
            Caption         =   "Place:"
            Height          =   225
            Left            =   90
            TabIndex        =   9
            Top             =   2985
            Width           =   2250
         End
         Begin VB.Label lblAdmin2 
            AutoSize        =   -1  'True
            Caption         =   "Admin5"
            Height          =   210
            Left            =   90
            TabIndex        =   8
            Top             =   675
            Width           =   2250
         End
      End
   End
End
Attribute VB_Name = "frmLocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private GIS As TatukGIS_XDK10.XGIS_Viewer
Private bPreventFlashEtc As Boolean
Private sSettlementCaption As String
Private sSettlementLayer As String
Private sSettlementUID As String
Private sSettlementField As String
Private dicUIDS0 As New Dictionary
Private dicUIDS1 As New Dictionary
Private dicUIDS2 As New Dictionary
Private dicUIDS3 As New Dictionary
Private dicUIDS4 As New Dictionary
Private dicUIDS5 As New Dictionary

Private Sub ComAdmin1_Click()
    ComAdmin2.Enabled = False
    ComAdmin3.Enabled = False
    ComAdmin4.Enabled = False
    ComAdmin5.Enabled = False
    ComAdmin6.Enabled = False
    ComAdmin2.Clear
    ComAdmin3.Clear
    ComAdmin4.Clear
    ComAdmin5.Clear
    ComAdmin6.Clear
    Call AdminClick(ComAdmin1, ComAdmin2, "AdminLevel0", "AdminLevel1", dicUIDS0, dicUIDS1)
End Sub

Private Sub ComAdmin2_Click()
    ComAdmin3.Enabled = False
    ComAdmin4.Enabled = False
    ComAdmin5.Enabled = False
    ComAdmin6.Enabled = False
    ComAdmin3.Clear
    ComAdmin4.Clear
    ComAdmin5.Clear
    ComAdmin6.Clear
    Call AdminClick(ComAdmin2, ComAdmin3, "AdminLevel1", "AdminLevel2", dicUIDS1, dicUIDS2)
End Sub

Private Sub ComAdmin3_Click()
    ComAdmin4.Enabled = False
    ComAdmin5.Enabled = False
    ComAdmin6.Enabled = False
    ComAdmin4.Clear
    ComAdmin5.Clear
    ComAdmin6.Clear
    Call AdminClick(ComAdmin3, ComAdmin4, "AdminLevel2", "AdminLevel3", dicUIDS2, dicUIDS3)
End Sub

Private Sub ComAdmin4_Click()
    ComAdmin5.Enabled = False
    ComAdmin6.Enabled = False
    ComAdmin5.Clear
    ComAdmin6.Clear
    Call AdminClick(ComAdmin4, ComAdmin5, "AdminLevel3", "AdminLevel4", dicUIDS3, dicUIDS4)
End Sub

Private Sub ComAdmin5_Click()
    ComAdmin6.Enabled = False
    ComAdmin6.Clear
    Call AdminClick(ComAdmin5, ComAdmin6, "AdminLevel4", "AdminLevel5", dicUIDS4, dicUIDS5)
End Sub

Private Sub ComAdmin6_Click()
    Call AdminClick(ComAdmin6, Nothing, "AdminLevel5", "", Nothing, Nothing)
End Sub

Public Sub Init(oGIS As TatukGIS_XDK10.XGIS_Viewer, _
                Optional sID As String)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        
        Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
        Dim bSettUsed As Boolean
        
100     bPreventFlashEtc = True
106     Set GIS = oGIS

108     ComAdmin2.Clear
110     ComAdmin1.Clear
112     ComAdmin6.Clear

        SafeMoveFirst g_RSAppSettings
        g_RSAppSettings.Find "SettingName = 'AdminLocation'"
        sSettlementCaption = g_RSAppSettings.Fields.Item("SettingValue4").value & ":"
        sSettlementLayer = g_RSAppSettings.Fields.Item("SettingValue1").value
        sSettlementUID = g_RSAppSettings.Fields.Item("SettingValue3").value
        sSettlementField = g_RSAppSettings.Fields.Item("SettingValue2").value
       
126     SafeMoveFirst g_RSAppSettings
128     g_RSAppSettings.Find "SettingName = 'AdminLevel1'"
130     lblAdmin2.caption = g_RSAppSettings.Fields.Item("SettingValue4").value & ":"

        If Not Len(lblAdmin2.caption) > 1 Then
            If Not bSettUsed And Len(ComAdmin1.Tag) < 2 Then
                ComAdmin2.Tag = "Settlement"
                lblAdmin2.caption = sSettlementCaption
                bSettUsed = True
            Else
                lblAdmin2.Visible = False
                ComAdmin2.Visible = False
                Me.Height = Me.Height - 500
            End If
        End If
        
136     SafeMoveFirst g_RSAppSettings
138     g_RSAppSettings.Find "SettingName = 'AdminLevel2'"
140     lblAdmin3.caption = g_RSAppSettings.Fields.Item("SettingValue4").value & ":"

        If Not Len(lblAdmin3.caption) > 1 Then
            If Not bSettUsed And Len(ComAdmin2.Tag) < 2 Then
                ComAdmin3.Tag = "Settlement"
                lblAdmin3.caption = sSettlementCaption
                bSettUsed = True
            Else
                lblAdmin3.Visible = False
                ComAdmin3.Visible = False
                Me.Height = Me.Height - 500
            End If

        End If
        
146     SafeMoveFirst g_RSAppSettings
148     g_RSAppSettings.Find "SettingName = 'AdminLevel3'"
150     lblAdmin4.caption = g_RSAppSettings.Fields.Item("SettingValue4").value & ":"

        If Not Len(lblAdmin4.caption) > 1 Then
            If Not bSettUsed And Len(ComAdmin3.Tag) < 2 Then
                ComAdmin4.Tag = "Settlement"
                lblAdmin4.caption = sSettlementCaption
                bSettUsed = True
            Else
                lblAdmin4.Visible = False
                ComAdmin4.Visible = False
                Me.Height = Me.Height - 500
            End If

        End If
    
156     SafeMoveFirst g_RSAppSettings
158     g_RSAppSettings.Find "SettingName = 'AdminLevel4'"
160     lblAdmin5.caption = g_RSAppSettings.Fields.Item("SettingValue4").value & ":"

        If Not Len(lblAdmin5.caption) > 1 Then
            If Not bSettUsed And Len(ComAdmin4.Tag) < 2 Then
                ComAdmin5.Tag = "Settlement"
                lblAdmin5.caption = sSettlementCaption
                bSettUsed = True
            Else
                lblAdmin5.Visible = False
                ComAdmin5.Visible = False
                Me.Height = Me.Height - 500
            End If

        End If
        
166     SafeMoveFirst g_RSAppSettings
168     g_RSAppSettings.Find "SettingName = 'AdminLevel5'"
170     lblAdmin6.caption = g_RSAppSettings.Fields.Item("SettingValue4").value & ":"

        If Not Len(lblAdmin6.caption) > 1 Then
            If Not bSettUsed And Len(ComAdmin5.Tag) < 2 Then
                ComAdmin6.Tag = "Settlement"
                lblAdmin6.caption = sSettlementCaption
                bSettUsed = True
            Else
            
                lblAdmin6.Visible = False
                ComAdmin6.Visible = False
                Me.Height = Me.Height - 500
            End If
       
        End If
        
176     SafeMoveFirst g_RSAppSettings
178     g_RSAppSettings.Find "SettingName = 'AdminLevel0'"
180     lblAdmin1.caption = g_RSAppSettings.Fields.Item("SettingValue4").value & ":"

        If Not Len(lblAdmin1.caption) > 1 Then
            
            lblAdmin1.Visible = False
            ComAdmin1.Visible = False
            Me.Height = Me.Height - 500
    
        End If

186     Set oLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").value)

188     If oLyr Is Nothing Then Exit Sub
        
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
        
        Set dicUIDS0 = New Dictionary
        
        For Each oShp9 In oLyr.Loop(oLyr.Extent, "", Nothing, "", True)
            ComAdmin1.AddItem oShp9.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value)
            DebugPrint oShp9.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value) & " = " & oShp9.uID
            dicUIDS0.Add oShp9.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value), oShp9.uID
        Next
        
200     If ComAdmin1.ListCount > 0 Then ComAdmin1.ListIndex = 0
202     bPreventFlashEtc = False
        
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLocator.Init " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
   
Private Sub AdminClick(COMClicked As ComboBox, _
                       COMNext As ComboBox, _
                       sPCODEClicked As String, _
                       sPCODENext As String, _
                       dicFrom As Dictionary, _
                       dicTo As Dictionary)
        
    Dim oLayerClicked As TatukGIS_XDK10.XGIS_LayerVector
    Dim oLayerNext As TatukGIS_XDK10.XGIS_LayerVector
    Dim oLayerFiltered As TatukGIS_XDK10.XGIS_LayerVector
    Dim oShapeClicked As TatukGIS_XDK10.XGIS_Shape
    Dim aShape As TatukGIS_XDK10.XGIS_Shape
    Dim lClickedShapeUID As Long
    Dim sLayerClicked As String
    Dim sFieldClicked As String
    Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
        
    If Not bPreventFlashEtc And Not COMClicked.List(COMClicked.ListIndex) = "" And Not IsNull(COMClicked.List(COMClicked.ListIndex)) And Not IsEmpty(COMClicked.List(COMClicked.ListIndex)) Then
             
        bPreventFlashEtc = True
        SafeMoveFirst g_RSAppSettings
        g_RSAppSettings.Find "SettingName = '" & IIf(Len(COMClicked.Tag) > 1, "AdminLocation", sPCODEClicked) & "'"
        
        sLayerClicked = IIf(Len(COMClicked.Tag) > 1, sSettlementLayer, g_RSAppSettings.Fields.Item("SettingValue1").value)
        Set oLayerClicked = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").value)

        If Not oLayerClicked Is Nothing Then
                    
            sFieldClicked = IIf(Len(COMClicked.Tag) > 1, sSettlementField, g_RSAppSettings.Fields.Item("SettingValue2").value)
            
            Set oShapeClicked = oLayerClicked.FindFirst(oLayerClicked.Extent, "GIS_UID = " & dicFrom.Item(COMClicked.List(COMClicked.ListIndex)), Nothing, "", True)
            
            If Not oShapeClicked Is Nothing Then

                lClickedShapeUID = oShapeClicked.uID

                If oShapeClicked.ShapeType = TatukGIS_XDK10.XgisShapeTypePoint Then
                    
                    'Dim oExtent As TatukGIS_XDK10.XGIS_Extent
                    'Set oExtent = New TatukGIS_XDK10.XGIS_Extent
                    'GIS.VisibleExtent = oShapeClicked.Extent
                    'oExtent.Prepare oShapeClicked.Extent.XMin + (GIS.VisibleExtent.XMax - GIS.VisibleExtent.XMin), GIS.VisibleExtent.YMin + (GIS.VisibleExtent.YMin - GIS.VisibleExtent.YMin), oShapeClicked.Extent.XMin + (GIS.VisibleExtent.XMax - GIS.VisibleExtent.XMin), GIS.VisibleExtent.YMin + (GIS.VisibleExtent.YMin - GIS.VisibleExtent.YMin)
                    'Set oShapeClicked = oLayerClicked.FindFirst(oLayerClicked.Extent, sFieldClicked & " = '" & COMClicked.List(COMClicked.ListIndex) & "'", Nothing, "", True)
                    GIS.CenterViewport oShapeClicked.Centroid
                    oShapeClicked.Flash 2, 50
                Else
                    GIS.VisibleExtent = oShapeClicked.Extent
                    oLayerClicked.GetShape(lClickedShapeUID).Flash 2, 50
                End If

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Not COMNext Is Nothing And Len(sPCODENext) > 1 Then
                        
                    SafeMoveFirst g_RSAppSettings

                    If Len(COMNext.Tag) > 1 Then
                        g_RSAppSettings.Find "SettingName = 'AdminLocation'"
                    Else
                        g_RSAppSettings.Find "SettingName = '" & sPCODENext & "'"
                    End If
                    
                    If Len(g_RSAppSettings.Fields.Item("SettingValue1").value) > 0 Then Set oLayerNext = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").value)
                    
                    If Not oLayerNext Is Nothing Then
                        Set dicTo = New Dictionary

                        For Each oShp9 In oLayerNext.Loop(oLayerNext.Extent, "", oShapeClicked, "T", True)
                            COMNext.AddItem oShp9.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value)
                            If Not dicTo.Exists(oShp9.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value)) Then dicTo.Add oShp9.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value), oShp9.uID
                        Next
    
                        If COMNext.ListCount > 0 Then COMNext.ListIndex = 0
                    End If
                        
                End If
                                                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End If

        End If
            
        If Not COMNext Is Nothing Then COMNext.Enabled = True
            
    End If
    
    bPreventFlashEtc = False

End Sub

