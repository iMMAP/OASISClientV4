VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Begin VB.Form frmW3Wizard 
   Caption         =   "Who What Where"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elSummary 
      Height          =   6705
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4185
      _cx             =   7382
      _cy             =   11827
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
      BorderWidth     =   2
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
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmW3Wizard.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin OASISClient.strWho strWho1 
         Height          =   4290
         Left            =   30
         TabIndex        =   3
         Top             =   2385
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   7567
      End
      Begin C1SizerLibCtl.C1Elastic elTop 
         Height          =   2355
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   4125
         _cx             =   7276
         _cy             =   4154
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
         Begin ActiveBar3LibraryCtl.ActiveBar3 AB 
            Height          =   2355
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   4125
            _LayoutVersion  =   2
            _ExtentX        =   7276
            _ExtentY        =   4154
            _DataPath       =   ""
            Bands           =   "frmW3Wizard.frx":0044
         End
      End
   End
End
Attribute VB_Name = "frmW3Wizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Public Event EditOrg(sCurrentOrg As String)
'Public Event EditOffice(sCurrentOff As String)
'
'Public Event MyOrganization()
'Public Event MyActivities()
'Public Event MyLocations()
'Public Event MyCluster()
'Private m_HasInitialized As Boolean
'Private m_iCurrentW3Module As Integer
'
'Public Function CurrentW3Module() As Integer
'        '<EhHeader>
'        On Error GoTo CurrentW3Module_Err
'        '</EhHeader>
'100     CurrentW3Module = m_iCurrentW3Module
'        '<EhFooter>
'        Exit Function
'
'CurrentW3Module_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmW3Wizard.CurrentW3Module " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Function
'
'Public Sub Init(CN As adodb.Connection)
'        '<EhHeader>
'        On Error GoTo Init_Err
'        '</EhHeader>
'100     strWho1.Init CN
'102     m_HasInitialized = True
'        '<EhFooter>
'        Exit Sub
'
'Init_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmW3Wizard.init " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Public Function HasInitialized() As Boolean
'        '<EhHeader>
'        On Error GoTo HasInitialized_Err
'        '</EhHeader>
'100     HasInitialized = m_HasInitialized
'        '<EhFooter>
'        Exit Function
'
'HasInitialized_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmW3Wizard.HasInitialized " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Function
'
'Private Sub AB_ChildBandChange(ByVal Band As ActiveBar3LibraryCtl.Band)
'        '<EhHeader>
'        On Error GoTo AB_ChildBandChange_Err
'        '</EhHeader>
'100     Select Case Band.Name
'
'            Case "cbContent"
'                'Where
'102             m_iCurrentW3Module = 2
'104             RaiseEvent MyLocations
'106         Case "cbOperations"
'                'What
'108             m_iCurrentW3Module = 1
'110             RaiseEvent MyActivities
'112         Case "cbProfile"
'                'Who
'114             m_iCurrentW3Module = 0
'116             RaiseEvent MyOrganization
'        End Select
'        'MsgBox "This functionality has not been implemented yet.", vbInformation, "OASIS DEVELOPMENT TEAM"
'        '<EhFooter>
'        Exit Sub
'
'AB_ChildBandChange_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmW3Wizard.AB_ChildBandChange " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub AB_ToolClick(ByVal Tool As ActiveBar3LibraryCtl.Tool)
'        '<EhHeader>
'        On Error GoTo AB_ToolClick_Err
'        '</EhHeader>
'100     MsgBox "This functionality has not been implemented yet.", vbInformation, "OASIS DEVELOPMENT TEAM"
'        '<EhFooter>
'        Exit Sub
'
'AB_ToolClick_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmW3Wizard.AB_ToolClick " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub Form_Load()
'If Not g_sLanguage = "" Then
'        If Not m_Cnn.State = adStateClosed Then
'            LoadLanguage Me.Name, g_sLanguage, m_Cnn
'        End If
'    End If
'End Sub
'
'Private Sub strWho1_Addlocation(sID As String)
'        'Dim cn As New ADODB.Connection
'        '<EhHeader>
'        On Error GoTo strWho1_Addlocation_Err
'        '</EhHeader>
'
'        'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OASIS\W3Import.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
'
'100     frmAddLocToW3.ctrWhere1.Init m_Cnn, False, sID, 1
'
'102     frmAddLocToW3.Show vbModal
'
'104     Unload frmAddLocToW3
'
'        '<EhFooter>
'        Exit Sub
'
'strWho1_Addlocation_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmW3Wizard.strWho1_Addlocation " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
