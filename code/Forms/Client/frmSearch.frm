VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3855
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elSearchMain 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3855
      _cx             =   6800
      _cy             =   11456
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
      GridRows        =   1
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmSearch.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1TSearchResults 
         Height          =   6495
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3855
         _cx             =   6800
         _cy             =   11456
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
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "Search|Results"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   0   'False
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Begin C1SizerLibCtl.C1Elastic elSearcgResults 
            Height          =   6120
            Left            =   4500
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   330
            Width           =   3765
            _cx             =   6641
            _cy             =   10795
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
            GridCols        =   5
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmSearch.frx":6885
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.CommandButton cmdCancel 
               Caption         =   "Cancel"
               Height          =   450
               Left            =   2130
               TabIndex        =   17
               Top             =   5580
               Width           =   1545
            End
            Begin VB.CommandButton cmdClearSearches 
               Caption         =   "Clear Searches"
               Height          =   450
               Left            =   90
               TabIndex        =   16
               Top             =   5580
               Width           =   1365
            End
            Begin vbalTreeViewLib6.vbalTreeView vbalSearch 
               Height          =   5430
               Left            =   90
               TabIndex        =   15
               Top             =   90
               Width           =   3585
               _ExtentX        =   6324
               _ExtentY        =   9578
               Indentation     =   30
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin C1SizerLibCtl.C1Elastic elSrchSettings 
            Height          =   6120
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   3765
            _cx             =   6641
            _cy             =   10795
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
            Begin VB.Frame Frame2 
               BorderStyle     =   0  'None
               Height          =   555
               Left            =   2220
               TabIndex        =   12
               Top             =   3060
               Width           =   1515
               Begin VB.CommandButton cmdOK 
                  Caption         =   "Search"
                  Height          =   405
                  Left            =   0
                  TabIndex        =   13
                  Top             =   0
                  Width           =   1455
               End
            End
            Begin VB.Frame FraSearch 
               Caption         =   "Settings:"
               Height          =   3015
               Left            =   0
               TabIndex        =   3
               Top             =   0
               Width           =   3855
               Begin VB.Frame FraLayerTo 
                  Caption         =   "Layer To Search:"
                  Height          =   1665
                  Left            =   60
                  TabIndex        =   7
                  Top             =   210
                  Width           =   3705
                  Begin VB.ComboBox ComLayer 
                     Height          =   315
                     Left            =   120
                     Style           =   2  'Dropdown List
                     TabIndex        =   11
                     Top             =   240
                     Width           =   3525
                  End
                  Begin VB.Frame FraFieldTo 
                     Caption         =   "Field To Search:"
                     Height          =   645
                     Left            =   60
                     TabIndex        =   9
                     Top             =   600
                     Width           =   3585
                     Begin VB.ComboBox ComFields 
                        Height          =   315
                        Left            =   120
                        Style           =   2  'Dropdown List
                        TabIndex        =   10
                        Top             =   210
                        Width           =   3375
                     End
                  End
                  Begin VB.CheckBox chkSearchAllFields 
                     Caption         =   "Search All Fields (May Be Slow)"
                     Height          =   285
                     Left            =   60
                     TabIndex        =   8
                     Top             =   1290
                     Width           =   3075
                  End
               End
               Begin VB.CheckBox chkSearchAll 
                  Caption         =   "Search All Layers (This is slow!)"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   6
                  Top             =   1920
                  Width           =   3435
               End
               Begin VB.Frame FraSearchValue 
                  Caption         =   "Search value:"
                  Height          =   615
                  Left            =   120
                  TabIndex        =   4
                  Top             =   2250
                  Width           =   3675
                  Begin VB.TextBox txtSearchVal 
                     Height          =   315
                     Left            =   60
                     TabIndex        =   5
                     Top             =   240
                     Width           =   3585
                  End
               End
            End
         End
      End
   End
   Begin vbalIml6.vbalImageList imgl 
      Left            =   0
      Top             =   2940
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   16
      Size            =   117096
      Images          =   "frmSearch.frx":68EE
      Version         =   131072
      KeyCount        =   102
      Keys            =   $"frmSearch.frx":23276
   End
   Begin VB.Menu mnuSearchRight 
      Caption         =   "mnuSearchRight"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveSearch 
         Caption         =   "Remove Search Item"
      End
   End
   Begin VB.Menu mnuGeoItem 
      Caption         =   "mnuGeoItem"
      Visible         =   0   'False
      Begin VB.Menu mnuSearchZoom 
         Caption         =   "Zoom To"
      End
      Begin VB.Menu mnuSearchFlash 
         Caption         =   "Flash"
      End
      Begin VB.Menu mnuSearchRemoveGeo 
         Caption         =   "Remove From Search"
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oShp As TatukGIS_XDK10.XGIS_Shape
Public LyrCol As Collection
Private m_bINIT As Boolean
Private m_OGIS As TatukGIS_XDK10.XGIS_Viewer
Dim m_fntItalic As StdFont
Dim m_bCancel As Boolean
Private m_sSearchCurUID As String
Private m_sSearchCurLyr As String
Private m_sSearchNodeKey As String

Public Sub Init(oGIS As TatukGIS_XDK10.XGIS_Viewer)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
    
        Dim i As Integer
        Dim sCurrItem As String
        
        m_bINIT = True
        Set m_OGIS = oGIS
100     Set LyrCol = New Collection
        
        On Error Resume Next
        
102     If ComLayer.ListCount > 0 Then
        
104         sCurrItem = ComLayer.List(ComLayer.ListIndex)
        
        End If
        
106     ComLayer.Clear
    
108     ComLayer.AddItem "--All--"
110     LyrCol.Add "--All--", "--All--"
        
        On Error Resume Next
        
112     For i = 0 To oGIS.Items.Count - 1
            
114         If GisUtils.IsInherited(oGIS.Items.Item(i), "XGIS_LayerVector") Then
116             LyrCol.Add oGIS.Items.Item(i).Name, oGIS.Items.Item(i).caption
118             ComLayer.AddItem oGIS.Items.Item(i).caption 'Name
            End If
        
        Next
       
120     If ComLayer.ListCount > 0 Then
            
122         If Len(sCurrItem) > 0 Then
124             FindIndexStrEx ComLayer, sCurrItem
            Else
126             FindIndexStrEx ComLayer, "--All--"
            End If
            
        End If
       
       ComFields.Clear
       ComFields.AddItem "--All--"
       ComFields.ListIndex = 0
       
       chkSearchAll.Value = vbUnchecked
       chkSearchAllFields.Value = vbUnchecked
       txtSearchVal.Text = ""
       
       m_bINIT = False
       
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSearch.Init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function AddSearchNode(sSearchID As String, sCaption As String) As cTreeViewNode
    vbalSearch.LabelEdit = True
    Set AddSearchNode = vbalSearch.Nodes.Add(, etvwChild, sSearchID, sCaption, imgl.ItemIndex(28))
    AddSearchNode.Bold = True
    AddSearchNode.ShowPlusMinus = True
End Function

Private Sub cmdCancel_Click()
    m_bCancel = True
End Sub

Private Sub cmdClearSearches_Click()
    PrepareTV
    'lstSearchResult.Clear
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer
    Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
    Dim sSQL As String
    Dim nodSearch As cTreeViewNode
    Dim nodValue As cTreeViewNode
    Dim nodResult As cTreeViewNode
    Dim k As Integer
    Dim j As Integer
    Dim sDel As String
    Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
    'lstSearchResult.Clear
    
    'TODO Strings works only now... Add the Doevents for this to cancel the Search. Do a nice way to present the found features and enable Zoom To!
    
    '  String = 0
    '  number = 1
    '  Float = 2
    '  Boolean = 3
    '  date = 4
    
    If chkSearchAllFields.Value = vbChecked Then

        With ComLayer
        
            If .List(.ListIndex) = "--All--" Then
            
                For j = 0 To m_OGIS.Items.Count - 1

                    If GisUtils.IsInherited(m_OGIS.Items.Item(j), "XGIS_LayerVector") Then
                        Set oLyr = m_OGIS.get(m_OGIS.Items.Item(j).Name)

                        If oLyr Is Nothing Then Exit Sub
                        
                        DoEvents
                        
                        If m_bCancel Then
                            m_bCancel = False
                            Exit Sub
                        End If
                        
                        Set nodSearch = AddSearchNode(oLyr.caption & ":::" & txtSearchVal.Text & ":::" & "--All--" & ":::" & Now, oLyr.caption & " : " & txtSearchVal.Text)

                        For i = 0 To oLyr.Fields.Count - 1
                            
                            Select Case oLyr.FieldInfo(CLng(i)).FieldType
                                Case TatukGIS_XDK10.XgisFieldTypeString
                                    sDel = "'"
                                Case TatukGIS_XDK10.XgisFieldTypeBoolean
                                    sDel = ""
                                Case TatukGIS_XDK10.XgisFieldTypeDate
                                    sDel = "#"
                                Case TatukGIS_XDK10.XgisFieldTypeFloat, TatukGIS_XDK10.XgisFieldTypeNumber
                                    sDel = ""
                            End Select
                            
                            If oLyr.FieldInfo(CLng(i)).FieldType = TatukGIS_XDK10.XgisFieldTypeString Then
                         
                                For Each oShp9 In oLyr.Loop(oLyr.Extent, oLyr.FieldInfo(CLng(i)).Name & " = '" & txtSearchVal.Text & "'", Nothing, "", True)                                    'lstSearchResult.AddItem oLyr.Shape.uID
                                    Set nodValue = nodSearch.AddChildNode(oShp9.uID & ":::" & Now() & Rnd(100), "GeoID:" & oShp9.uID, imgl.ItemIndex(97)) ', imgl.ItemIndex(99))
                                    nodValue.ShowPlusMinus = True
                                    nodValue.Tag = oLyr.Name
                                    
                                    For k = 0 To oLyr.Fields.Count - 1
                        
                                        Set nodResult = nodValue.AddChildNode(oShp9.uID & ":::" & Now() & Rnd(100), oLyr.Fields.Item(k).Name & " = " & oShp9.GetField(oLyr.Fields.Item(k).Name), imgl.ItemIndex(64))
                                        
                                        If i = k Then nodResult.Bold = True

                                        DoEvents
                                        
                                        If m_bCancel Then
                                            m_bCancel = False
                                            Exit Sub
                                        End If

                                    Next

                                Next
              
                            End If

                        Next

                    End If
                    
                    If Not nodSearch Is Nothing Then
                        If nodSearch.Children.Count = 0 Then
                            vbalSearch.Nodes.Remove nodSearch.Key
                        End If
                    End If
                    
                Next j

            Else
                Set oLyr = m_OGIS.get(LyrCol.Item(ComLayer.List(ComLayer.ListIndex)))

                If oLyr Is Nothing Then Exit Sub
                
                Set nodSearch = AddSearchNode(oLyr.caption & ":::" & txtSearchVal.Text & ":::" & "--All--" & ":::" & Now, oLyr.caption & " : " & txtSearchVal.Text)

                For i = 0 To oLyr.Fields.Count - 1

                    If oLyr.FieldInfo(CLng(i)).FieldType = TatukGIS_XDK10.XgisFieldTypeString Then
                        
                        For Each oShp9 In oLyr.Loop(oLyr.Extent, oLyr.FieldInfo(CLng(i)).Name & " = '" & txtSearchVal.Text & "'", Nothing, "", True)
                                          
                            Set nodValue = nodSearch.AddChildNode(oShp9.uID & ":::" & Now() & Rnd(100), "GeoID:" & oShp9.uID, imgl.ItemIndex(97)) ', imgl.ItemIndex(99))
                            nodValue.ShowPlusMinus = True
                            nodValue.Tag = oLyr.Name
                            For k = 0 To oLyr.Fields.Count - 1
                        
                                Set nodResult = nodValue.AddChildNode(oShp9.uID & ":::" & Now() & Rnd(100), oLyr.Fields.Item(k).Name & " = " & oShp9.GetField(oLyr.Fields.Item(k).Name), imgl.ItemIndex(64))

                                If i = k Then nodResult.Bold = True
                            Next

                        Next
              
                    End If

                Next
              
            End If
        
        End With
                
    ElseIf chkSearchAll.Value = vbChecked Then
    
    Else
        
        Set oLyr = m_OGIS.get(LyrCol.Item(ComLayer.List(ComLayer.ListIndex)))
        
        If oLyr Is Nothing Then Exit Sub
        
        'Set nodSearch = AddSearchNode(oLyr.caption & ":::" & txtSearchVal.Text & ":::" & Now, oLyr.caption & " : " & txtSearchVal.Text)
        
        If ComFields.List(ComFields.ListIndex) = "--All--" Then
            Set nodSearch = AddSearchNode(oLyr.caption & ":::" & txtSearchVal.Text & ":::" & "--All--" & ":::" & Now, oLyr.caption & " : " & txtSearchVal.Text)
                    
            If oLyr.FieldInfo(oLyr.FindField(ComFields.List(ComFields.ListIndex))).FieldType = TatukGIS_XDK10.XgisFieldTypeString Then
                
                For k = 0 To oLyr.Field.Count - 1
                    For Each oShp9 In oLyr.Loop(oLyr.Extent, ComFields.List(ComFields.ListIndex) & " = '" & txtSearchVal.Text & "'", Nothing, "", True)
                         Set nodValue = nodSearch.AddChildNode(oShp9.uID & ":::" & Now(), "GeoID:" & oShp9.uID, imgl.ItemIndex(97)) ', imgl.ItemIndex(99))
                        nodValue.ShowPlusMinus = True
                        nodValue.Tag = oLyr.Name
                        
                        For i = 0 To oLyr.Field.Count - 1
                    
                            nodValue.AddChildNode oShp9.uID & ":::" & Now() & Rnd(100), oLyr.Fields.Item(i).Name & " = " & oShp9.GetField(oLyr.Fields.Item(i).Name), imgl.ItemIndex(64)
                        Next

                    Next

                Next

            End If
                    
        Else
            Set nodSearch = AddSearchNode(oLyr.caption & ":::" & txtSearchVal.Text & ":::" & ComFields.List(ComFields.ListIndex) & ":::" & Now, "Search: " & oLyr.caption & " for " & txtSearchVal.Text & " in " & ComFields.List(ComFields.ListIndex))

            If oLyr.FieldInfo(oLyr.FindField(ComFields.List(ComFields.ListIndex))).FieldType = TatukGIS_XDK10.XgisFieldTypeString Then
                
                For Each oShp9 In oLyr.Loop(oLyr.Extent, ComFields.List(ComFields.ListIndex) & " = '" & txtSearchVal.Text & "'", Nothing, "", True)
                    'lstSearchResult.AddItem oLyr.Shape.GetField(ComFields.List(ComFields.ListIndex))
                    Set nodValue = nodSearch.AddChildNode(oShp9.uID & ":::" & Now(), "GeoID:" & oShp9.uID, imgl.ItemIndex(97)) ', imgl.ItemIndex(99))
                    nodValue.ShowPlusMinus = True
                    nodValue.Tag = oLyr.Name
                    
                    For i = 0 To oLyr.Field.Count - 1
                        nodValue.AddChildNode oShp9.uID & ":::" & Now() & Rnd(100), oLyr.Fields.Item(i).Name & " = " & oShp9.GetField(oLyr.Fields.Item(i).Name), imgl.ItemIndex(64)
                    Next
                Next
              
            End If
              
        End If
        
    End If

End Sub

Private Sub ComLayer_Click()
        '<EhHeader>
        On Error GoTo ComLayer_Click_Err
        '</EhHeader>
        Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
        Dim i As Integer
100     If m_bINIT Then Exit Sub
    
102     'DebugPrint
    
104     Set oLyr = m_OGIS.get(LyrCol.Item(ComLayer.List(ComLayer.ListIndex)))
    
106     If oLyr Is Nothing Then Exit Sub
    
108     With oLyr.Fields
        
            
110         ComFields.Clear
        
112         For i = 0 To .Count - 1
114             ComFields.AddItem .Item(i).Name
                DebugPrint .Item(i).Name
                DebugPrint oLyr.FieldInfo(i).Binary
                DebugPrint oLyr.FieldInfo(i).Decimal
                DebugPrint oLyr.FieldInfo(i).Deleted
                DebugPrint oLyr.FieldInfo(i).ExportName
                DebugPrint oLyr.FieldInfo(i).FieldType
                DebugPrint oLyr.FieldInfo(i).FileFormat
                DebugPrint oLyr.FieldInfo(i).Hidden
                DebugPrint oLyr.FieldInfo(i).Predefied
                DebugPrint oLyr.FieldInfo(i).Width
            Next
        
        End With
    
        '<EhFooter>
        Exit Sub

ComLayer_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSearch.ComLayer_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
    PrepareTV
End Sub

Public Sub PrepareTV()
    With vbalSearch
        .ImageList = imgl.hIml
        .Nodes.Clear
    End With
    
    Set m_fntItalic = New StdFont
    m_fntItalic.Name = "Tahoma"
    m_fntItalic.Size = 8.25
    'm_fntItalic.Italic = True
    m_fntItalic.Underline = True
    m_fntItalic.Bold = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_OGIS = Nothing
End Sub

Private Sub mnuRemoveSearch_Click()
    If m_sSearchNodeKey = "" Then Exit Sub
    vbalSearch.Nodes.Remove m_sSearchNodeKey
    m_sSearchNodeKey = ""
End Sub

Private Sub mnuSearchFlash_Click()
    If m_sSearchCurUID = "" Then Exit Sub
    
    Dim lShape As TatukGIS_XDK10.XGIS_Shape
        
    Set lShape = m_OGIS.get(m_sSearchCurLyr).GetShape(CLng(m_sSearchCurUID))
    lShape.Flash
        
    m_sSearchCurLyr = ""
    m_sSearchCurUID = ""
    
    Set lShape = Nothing
End Sub

Private Sub mnuSearchRemoveGeo_Click()
    If m_sSearchNodeKey = "" Then Exit Sub
    vbalSearch.Nodes.Remove m_sSearchNodeKey
    m_sSearchNodeKey = ""
End Sub

Private Sub mnuSearchZoom_Click()

    If m_sSearchCurUID = "" Then Exit Sub
    
    Dim i As Integer
    Dim lShape As TatukGIS_XDK10.XGIS_Shape
        
    Set lShape = m_OGIS.get(m_sSearchCurLyr).GetShape(CLng(m_sSearchCurUID))

    m_OGIS.Lock
    m_OGIS.VisibleExtent = lShape.Extent
    m_OGIS.Unlock
        
    m_sSearchCurLyr = ""
    m_sSearchCurUID = ""
    
    Set lShape = Nothing
End Sub

Private Sub vbalSearch_NodeRightClick(Node As vbalTreeViewLib6.cTreeViewNode)
    Dim tp As POINTAPI
    Dim sKey() As String
    
    If Node.Children.Count = 0 Then Exit Sub 'Attribute field
   
    m_sSearchNodeKey = Node.Key
   
    GetCursorPos tp
    ScreenToClient vbalSearch.hWnd, tp
   
    If Node.Parent Is Nothing Then
        DebugPrint "Root"
        Me.PopupMenu mnuSearchRight, , vbalSearch.Left + tp.x * Screen.TwipsPerPixelX, vbalSearch.Top + tp.y * Screen.TwipsPerPixelY
   
    ElseIf Node.Parent.Parent Is Nothing Then
        DebugPrint "Geoitem"
        m_sSearchCurLyr = Node.Tag
        sKey = Split(Node.Key, ":::")
        m_sSearchCurUID = sKey(0)
        Me.PopupMenu mnuGeoItem, , vbalSearch.Left + tp.x * Screen.TwipsPerPixelX, vbalSearch.Top + tp.y * Screen.TwipsPerPixelY
   
    End If
   
End Sub
