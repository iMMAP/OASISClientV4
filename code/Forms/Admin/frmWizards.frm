VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Begin VB.Form frmWizards 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wizards"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7695
   Icon            =   "frmWizards.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7695
   StartUpPosition =   1  'CenterOwner
   Begin C1SizerLibCtl.C1Elastic c1Utils 
      Height          =   3945
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   7695
      _cx             =   13573
      _cy             =   6959
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
      Begin CONTROLSLibCtl.dxPicBtn cmdSynchLayers 
         Height          =   720
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmWizards.frx":6852
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn dxGeoBookmarks 
         Height          =   720
         Left            =   2400
         TabIndex        =   3
         Top             =   330
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmWizards.frx":75A4
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdGISAttrib 
         Height          =   720
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmWizards.frx":82F6
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdEncryption 
         Height          =   720
         Left            =   2400
         TabIndex        =   7
         Top             =   1500
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmWizards.frx":9048
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdMapProducts 
         Height          =   720
         Left            =   4440
         TabIndex        =   9
         Top             =   360
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmWizards.frx":9D9A
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdFeeds 
         Height          =   720
         Left            =   4440
         TabIndex        =   11
         Top             =   1560
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmWizards.frx":AAEC
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdThemes 
         Height          =   720
         Left            =   360
         TabIndex        =   13
         Top             =   2820
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmWizards.frx":B83E
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdDynamicDataDefs 
         Height          =   720
         Left            =   2400
         TabIndex        =   15
         Top             =   2730
         Visible         =   0   'False
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmWizards.frx":C590
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdDynamicReports 
         Height          =   720
         Left            =   4440
         TabIndex        =   17
         Top             =   2730
         Visible         =   0   'False
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmWizards.frx":D2E2
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdMapPrintTemplates 
         Height          =   720
         Left            =   6420
         TabIndex        =   19
         Top             =   330
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmWizards.frx":E034
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdChartGenerator 
         Height          =   720
         Left            =   6390
         TabIndex        =   21
         Top             =   1530
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmWizards.frx":ED86
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Chart Generator"
         Height          =   375
         Left            =   6000
         TabIndex        =   22
         Top             =   2310
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Map Print Templates"
         Height          =   375
         Left            =   6030
         TabIndex        =   20
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Dynamic Reports"
         Height          =   375
         Left            =   4050
         TabIndex        =   18
         Top             =   3510
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Dynamic Data Defs"
         Height          =   375
         Left            =   2010
         TabIndex        =   16
         Top             =   3510
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Themes"
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   3540
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Feeds"
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Map Products"
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Encryption Wizard"
         Height          =   375
         Left            =   2070
         TabIndex        =   8
         Top             =   2250
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "GIS Attributes"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Synch Layers"
         Height          =   375
         Left            =   0
         TabIndex        =   0
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "GeoMarks Explorer"
         Height          =   375
         Left            =   2010
         TabIndex        =   4
         Top             =   1110
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmWizards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_frmSelectUserGroup As frmSelectUserGroup
Dim m_frmSynchLayerWizard As frmSynchLayerWizard
Dim m_frmGeoMarksExplorer As frmGeoMarksExplorer
Dim m_frmEncryptWizard As frmEncryptWizard
Dim m_frmMapProductsWiz As frmMapProductsWiz
Dim m_frmGISAttrWiz As frmGISAttrWiz
Dim m_frmFeedsWizard As frmFeedsWizard
Dim m_frmThemeWiz As frmThemeWiz

Dim RSLocalUserGroups As New ADODB.Recordset

'Dim m_frmDynamicReports As frmDynamicReports
Dim m_frmChartWiz As frmChartWiz
Dim m_frmMapPrint As frmMapPrint
Dim m_frmDynamicDataMenu As frmDynamicDataMenu

Public Sub setUserGroupsRS(ByRef PassedRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo setUserGroupsRS_Err
        '</EhHeader>
        
100     Set RSLocalUserGroups = PassedRS

        '<EhFooter>
        Exit Sub

setUserGroupsRS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.setUserGroupsRS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdChartGenerator_Click()
        '<EhHeader>
        On Error GoTo cmdChartGenerator_Click_Err
        '</EhHeader>
100     If Not m_frmChartWiz.Visible Then
102         m_frmChartWiz.Show vbModeless, Me
        End If

        '<EhFooter>
        Exit Sub

cmdChartGenerator_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.cmdChartGenerator_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdDynamicReports_Click()
        '<EhHeader>
        On Error GoTo cmdDynamicReports_Click_Err
        '</EhHeader>
100     If Not m_frmDynamicReports.Visible Then
102         m_frmDynamicReports.Init False
104         m_frmDynamicReports.Show vbModeless, Me
        End If

        '<EhFooter>
        Exit Sub

cmdDynamicReports_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.cmdDynamicReports_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdMapPrintTemplates_Click()
        '<EhHeader>
        On Error GoTo cmdMapPrintTemplates_Click_Err
        '</EhHeader>
100     If Not m_frmSelectUserGroup.Visible Then
102         Set m_frmSelectUserGroup.dxUG.DataSource = RSLocalUserGroups
104         m_frmSelectUserGroup.dxUG.Columns.RetrieveFields
106         m_frmSelectUserGroup.dxUG.Dataset.Refresh
108         m_frmSelectUserGroup.Show vbModal

110         If m_frmSelectUserGroup.Tag = True Then
112             m_frmMapPrint.Init WebSite, RSLocalUserGroups!Name
114             m_frmMapPrint.Show vbModeless, Me
            End If
        End If

        '<EhFooter>
        Exit Sub

cmdMapPrintTemplates_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.cmdMapPrintTemplates_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdDynamicDataDefs_Click()
        '<EhHeader>
        On Error GoTo cmdDynamicDataDefs_Click_Err
        '</EhHeader>
100     If Not m_frmSelectUserGroup.Visible Then
102         Set m_frmSelectUserGroup.dxUG.DataSource = RSLocalUserGroups
104         m_frmSelectUserGroup.dxUG.Columns.RetrieveFields
106         m_frmSelectUserGroup.dxUG.Dataset.Refresh
108         m_frmSelectUserGroup.Show vbModal
    
110         If m_frmSelectUserGroup.Tag = True Then
112             m_frmDynamicDataMenu.setUserGroupsRS RSLocalUserGroups!Name, RSLocalUserGroups!sGUID
114             m_frmDynamicDataMenu.Init WebSite
116             m_frmDynamicDataMenu.Show vbModeless, Me
            End If
        End If

        '<EhFooter>
        Exit Sub

cmdDynamicDataDefs_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.cmdDynamicDataDefs_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdFeeds_Click()
        '<EhHeader>
        On Error GoTo cmdFeeds_Click_Err
        '</EhHeader>
100     If Not m_frmSelectUserGroup.Visible Then
102         Set m_frmSelectUserGroup.dxUG.DataSource = RSLocalUserGroups
104         m_frmSelectUserGroup.dxUG.Columns.RetrieveFields
106         m_frmSelectUserGroup.dxUG.Dataset.Refresh
108         m_frmSelectUserGroup.Show vbModal

110         If m_frmSelectUserGroup.Tag = True Then
112             m_frmFeedsWizard.setUserGroupsRS RSLocalUserGroups
114             m_frmFeedsWizard.Show vbModeless, Me
            End If
        End If

        '<EhFooter>
        Exit Sub

cmdFeeds_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.cmdFeeds_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdGISAttrib_Click()
        '<EhHeader>
        On Error GoTo cmdGISAttrib_Click_Err
        '</EhHeader>
100     If Not m_frmSelectUserGroup.Visible Then
102         Set m_frmSelectUserGroup.dxUG.DataSource = RSLocalUserGroups
104         m_frmSelectUserGroup.dxUG.Columns.RetrieveFields
106         m_frmSelectUserGroup.dxUG.Dataset.Refresh
108         m_frmSelectUserGroup.Show vbModal

110         If m_frmSelectUserGroup.Tag = True Then
112             m_frmGISAttrWiz.setUserGroupsRS RSLocalUserGroups
114             m_frmGISAttrWiz.Show vbModeless, Me
            End If
        End If

        '<EhFooter>
        Exit Sub

cmdGISAttrib_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.cmdGISAttrib_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdMapProducts_Click()
        '<EhHeader>
        On Error GoTo cmdMapProducts_Click_Err
        '</EhHeader>
100     If Not m_frmSelectUserGroup.Visible Then
102         Set m_frmSelectUserGroup.dxUG.DataSource = RSLocalUserGroups
104         m_frmSelectUserGroup.dxUG.Columns.RetrieveFields
106         m_frmSelectUserGroup.dxUG.Dataset.Refresh
108         m_frmSelectUserGroup.Show vbModal

110         If m_frmSelectUserGroup.Tag = True Then
112             m_frmMapProductsWiz.setUserGroupsRS RSLocalUserGroups
114             m_frmMapProductsWiz.Show vbModeless, Me
            End If
        End If

        '<EhFooter>
        Exit Sub

cmdMapProducts_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.cmdMapProducts_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSynchLayers_Click()
        '<EhHeader>
        On Error GoTo cmdSynchLayers_Click_Err
        '</EhHeader>
100     If Not m_frmSelectUserGroup.Visible Then
102         Set m_frmSelectUserGroup.dxUG.DataSource = RSLocalUserGroups
104         m_frmSelectUserGroup.dxUG.Columns.RetrieveFields
106         m_frmSelectUserGroup.dxUG.Dataset.Refresh
108         m_frmSelectUserGroup.Show vbModal

110         If m_frmSelectUserGroup.Tag = True Then
112             m_frmSynchLayerWizard.setUserGroupsRS RSLocalUserGroups
114             m_frmSynchLayerWizard.Show vbModeless, Me
            End If
        End If

        '<EhFooter>
        Exit Sub

cmdSynchLayers_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.cmdSynchLayers_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdThemes_Click()
        '<EhHeader>
        On Error GoTo cmdThemes_Click_Err
        '</EhHeader>
100     If Not m_frmSelectUserGroup.Visible Then
102         Set m_frmSelectUserGroup.dxUG.DataSource = RSLocalUserGroups
104         m_frmSelectUserGroup.dxUG.Columns.RetrieveFields
106         m_frmSelectUserGroup.dxUG.Dataset.Refresh
108         m_frmSelectUserGroup.Show vbModal

110         If m_frmSelectUserGroup.Tag = True Then
112             m_frmThemeWiz.setUserGroupsRS RSLocalUserGroups
114             m_frmThemeWiz.Show vbModeless, Me
            End If
        End If

        '<EhFooter>
        Exit Sub

cmdThemes_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.cmdThemes_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxGeoBookmarks_Click()
        '<EhHeader>
        On Error GoTo dxGeoBookmarks_Click_Err
        '</EhHeader>
100     If Not m_frmSelectUserGroup.Visible Then
102         Set m_frmSelectUserGroup.dxUG.DataSource = RSLocalUserGroups
104         m_frmSelectUserGroup.dxUG.Columns.RetrieveFields
106         m_frmSelectUserGroup.dxUG.Dataset.Refresh
108         m_frmSelectUserGroup.Show vbModal

110         If m_frmSelectUserGroup.Tag = True Then
112             m_frmGeoMarksExplorer.setUserGroupsRS RSLocalUserGroups
114             m_frmGeoMarksExplorer.Show vbModeless, Me
            End If
        End If

        '<EhFooter>
        Exit Sub

dxGeoBookmarks_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.dxGeoBookmarks_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdEncryption_Click()
        '<EhHeader>
        On Error GoTo cmdEncryption_Click_Err
        '</EhHeader>
100     If Not m_frmEncryptWizard.Visible Then
102         m_frmEncryptWizard.Show vbModeless, Me
        End If

        '<EhFooter>
        Exit Sub

cmdEncryption_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.cmdEncryption_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     Set m_frmChartWiz = New frmChartWiz
102    ' Set m_frmDynamicReports = New frmDynamicReports
104     Set m_frmMapPrint = New frmMapPrint
106     Set m_frmDynamicDataMenu = New frmDynamicDataMenu
108     Set m_frmSelectUserGroup = New frmSelectUserGroup
110     Set m_frmFeedsWizard = New frmFeedsWizard
112     Set m_frmGISAttrWiz = New frmGISAttrWiz
114     Set m_frmMapProductsWiz = New frmMapProductsWiz
116     Set m_frmSynchLayerWizard = New frmSynchLayerWizard
118     Set m_frmThemeWiz = New frmThemeWiz
120     Set m_frmGeoMarksExplorer = New frmGeoMarksExplorer
122     Set m_frmEncryptWizard = New frmEncryptWizard
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

        Unload m_frmSelectUserGroup

100     Set m_frmSelectUserGroup = Nothing
102     Set m_frmSynchLayerWizard = Nothing
104     Set m_frmGeoMarksExplorer = Nothing
106     Set m_frmGISAttrWiz = Nothing
108     Set m_frmEncryptWizard = Nothing
110     Set m_frmFeedsWizard = Nothing
112     Set m_frmMapProductsWiz = Nothing
114     Set m_frmDynamicReports = Nothing
116     Set m_frmChartWiz = Nothing
118     Set m_frmMapPrint = Nothing
120     Set m_frmDynamicDataMenu = Nothing
122     Set m_frmThemeWiz = Nothing

      '  On Error Resume Next

124     'For i = 1 To Forms.Count

126        ' If Forms(i).Name <> Me.Name Then
128         '    Unload Forms(i)
130         '    Set Forms(i) = Nothing
           ' End If

       ' Next

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmWizards.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

