VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{605925BE-4799-4093-A2E7-39323147E70E}#1.0#0"; "C1Query8.OCX"
Begin VB.Form frmOASISChartOCTfiles 
   Caption         =   "Select Chart Template to display"
   ClientHeight    =   4620
   ClientLeft      =   315
   ClientTop       =   615
   ClientWidth     =   4785
   Icon            =   "frmOASISChartOCTFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4785
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   4620
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4785
      _cx             =   8440
      _cy             =   8149
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
      BackColor       =   5292196
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   0
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
      GridRows        =   12
      GridCols        =   12
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmOASISChartOCTFiles.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   4620
         Left            =   1980
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   2805
         _cx             =   4948
         _cy             =   8149
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
         BackColor       =   5292196
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
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmOASISChartOCTFiles.frx":698C
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.FileListBox File1 
            Height          =   4380
            Left            =   90
            Pattern         =   "*.oct"
            TabIndex        =   3
            Top             =   90
            Width           =   2625
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Draw Chart"
         Height          =   420
         Left            =   255
         TabIndex        =   1
         Top             =   3810
         Width           =   1410
      End
   End
   Begin C1Query80Ctl.C1QueryFrame C1QueryFrame1 
      Height          =   2115
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      _cx             =   7646
      _cy             =   3731
      DesignTemplates =   ""
      ManualRender    =   0   'False
      Enabled         =   -1  'True
      DebugContextMenu=   0   'False
      Border          =   -1  'True
      TabInQuery      =   0   'False
      FullFieldNames  =   0   'False
      SchemaControl   =   "C1Query1"
      ContentsType    =   2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DesignTimeTemplates=   -1  'True
      TypedEditing    =   -1  'True
      FormatDate      =   2
      CheckBoxes      =   0   'False
      CheckValues     =   -1  'True
   End
   Begin C1Query80Ctl.C1Query C1Query1 
      Height          =   540
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   540
      _cx             =   952
      _cy             =   952
      DesignTemplates =   ""
      MainViewName    =   ""
      DataMember      =   ""
      FilterMode      =   0   'False
      ApplyExtensions =   3
      NameSubstitute  =   ""
      SaveSchemaAsString=   0   'False
      PathSeparator   =   "."
      SchemaData      =   "frmOASISChartOCTFiles.frx":69C4
   End
End
Attribute VB_Name = "frmOASISChartOCTfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_frmOASISCharts As frmOASISCharts

Private Function FileExists(sFullPath As String) As Boolean
        '<EhHeader>
        On Error GoTo FileExists_Err
        '</EhHeader>

        Dim oFile As New Scripting.FileSystemObject
100     FileExists = oFile.FileExists(sFullPath)

        '<EhFooter>
        Exit Function

FileExists_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISChartOCTfiles.FileExists " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub Command1_Click()
        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>
    
        Dim udtOASISChartO As OASISChartObj
    
100     If FileExists(File1.Path & "\" & File1.Filename) Then
102         Set m_frmOASISCharts = New frmOASISCharts

            If FileExists(File1.Path & "\SQL_" & Replace$(File1.Filename, ".oct", ".xml")) Then
                C1QueryFrame1.LoadFromXMLFile File1.Path & "\SQL_" & Replace$(File1.Filename, ".oct", ".xml")
                C1QueryFrame1.Render
            End If

104         With udtOASISChartO

106             With .udtChartTemplate
108                 .enmFormat = tplBin
                    '.sDecription = "Some Potato Junkie"
110                 .sName = File1.Path & "\" & File1.Filename
                End With
                
                .sSQL = C1Query1.SQL
            End With
                
'112         With udtOASISChartO
'114             .bAnnoTBR = True
'116             .bChartTBR = True
'            End With
                
118         m_frmOASISCharts.SetChart udtOASISChartO
120         m_frmOASISCharts.Show vbModeless, Me
    
122         Set m_frmOASISCharts = Nothing
    
        Else
        
124         MsgBox "Select a template file!"
    
        End If
    
        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISChartOCTfiles.Command1_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     Me.File1.Path = g_sAppPath & "\data\templates\ChartTemplates"
'102     C1Elastic1.Picture = g_PictureDialogSmall
104     C1Elastic1.PicturePos = ppLeftTop
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISChartOCTfiles.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     Set m_frmOASISCharts = Nothing
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISChartOCTfiles.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
