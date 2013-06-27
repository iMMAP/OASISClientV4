VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F7E0465A-9B48-4A2A-9144-10D6AFAB4BBB}#1.0#0"; "OASISDynamReports.ocx"
Begin VB.Form frmDynamicReports 
   Caption         =   "Dynamic Reporting"
   ClientHeight    =   6450
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   7485
   Icon            =   "frmDynamicReports.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   7485
   StartUpPosition =   1  'CenterOwner
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6450
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7485
      _cx             =   13203
      _cy             =   11377
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
      _GridInfo       =   $"frmDynamicReports.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin OASISDynamicReports.OASISDynamReports OASISDynamReports 
         Height          =   6270
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   11060
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu File_Open 
         Caption         =   "Open"
      End
      Begin VB.Menu File_Save 
         Caption         =   "Save"
      End
      Begin VB.Menu File_Print 
         Caption         =   "Print"
         Visible         =   0   'False
      End
      Begin VB.Menu File_Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu Page 
      Caption         =   "Page"
      Begin VB.Menu Page_Portrait 
         Caption         =   "Portrait"
      End
      Begin VB.Menu Page_Landscape 
         Caption         =   "Landscape"
      End
   End
   Begin VB.Menu Captions 
      Caption         =   "Captions"
      Begin VB.Menu Captions_Title 
         Caption         =   "Title"
      End
   End
   Begin VB.Menu Experimental 
      Caption         =   "Experimental"
      Visible         =   0   'False
      Begin VB.Menu Experimental_GenCharts 
         Caption         =   "Generate Charts"
      End
      Begin VB.Menu Experimental_ShowReportDetail 
         Caption         =   "Show Report Detail"
      End
      Begin VB.Menu Test 
         Caption         =   "Test"
      End
   End
End
Attribute VB_Name = "frmDynamicReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ss As String
Dim bIsAdminLoggedIn As String
Public sDynamDBPath As String

Private Sub Captions_Title_Click()
        '<EhHeader>
        On Error GoTo Captions_Title_Click_Err
        '</EhHeader>
100     Me.OASISDynamReports.SetTitle
        '<EhFooter>
        Exit Sub

Captions_Title_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicReports.Captions_Title_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub File_Close_Click()
        '<EhHeader>
        On Error GoTo File_Close_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

File_Close_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicReports.File_Close_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub File_Open_Click()
        '<EhHeader>
        On Error GoTo File_Open_Click_Err
        '</EhHeader>

        Dim c As New cCommonDialog
        On Error Resume Next
        'CreateDynamDBPath
    
100     c.DefaultExt = "*.xml"
102     c.DialogTitle = "Open Report Definitions"
104     c.Filter = "Report Definitions (*.xml)|*.xml"
        'c.InitDir = "%temp%" 'sDynamDBPath
106     c.InitDir = sDynamDBPath
  
108     c.ShowOpen
    
110     If Not c.Filename = "" Then
        
112         Me.OASISDynamReports.loadXML c.Filename
    
        End If
    
114     Me.Page_Portrait.Checked = Not Me.OASISDynamReports.GetOrientIsPortrait
116     Me.Page_Landscape.Checked = Me.OASISDynamReports.GetOrientIsPortrait

        '<EhFooter>
        Exit Sub

File_Open_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicReports.File_Open_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub CreateDynamDBPath()
        '<EhHeader>
        On Error GoTo CreateDynamDBPath_Err
        '</EhHeader>

100     If bIsAdminLoggedIn Then
102         sDynamDBPath = CreateAppPath & "\Data\db\dynamicdata"
        Else
104         sDynamDBPath = g_sAppPath & "\Data\db\dynamicdata"
        End If

        'Stop
        '<EhFooter>
        Exit Sub

CreateDynamDBPath_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicReports.CreateDynamDBPath " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Init(bAmIClient As Boolean, _
                Optional sConnString As String, _
                Optional sXMLFile As String)
        '<EhHeader>
        On Error GoTo init_Err
        '</EhHeader>
100     bIsAdminLoggedIn = Not bAmIClient
        
102     If Len(sConnString) < 1 Then
104         CreateDynamDBPath
        End If
    
106     If Len(sXMLFile) < 1 Then
108         File_Open_Click
        Else
110         Me.OASISDynamReports.loadXML sXMLFile
112         Me.Page_Portrait.Checked = Not Me.OASISDynamReports.GetOrientIsPortrait
114         Me.Page_Landscape.Checked = Me.OASISDynamReports.GetOrientIsPortrait
        End If
    
        '<EhFooter>
        Exit Sub

init_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicReports.Init " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub File_Save_Click()
    
    Dim c As New cCommonDialog
    'On Error Resume Next

    c.DefaultExt = "*.pdf"
    c.DialogTitle = "Save to PDF"
    c.Filter = "PDF File (*.pdf)|*.pdf"
    c.InitDir = "%userprofile%"
    c.ShowSave
    
    If Not c.Filename = "" Then
        
        Me.OASISDynamReports.PrintToPDF c.Filename, sDynamDBPath
    
    End If
    
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>
100     If Not g_sLanguage = "" Then
102         If Not m_Cnn.State = adStateClosed Then
104             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicReports.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Page_Landscape_Click()
        '<EhHeader>
        On Error GoTo Page_Landscape_Click_Err
        '</EhHeader>
100     Me.OASISDynamReports.SetOrientLandscape
102     Me.Page_Portrait.Checked = False
104     Me.Page_Landscape.Checked = True
        '<EhFooter>
        Exit Sub

Page_Landscape_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicReports.Page_Landscape_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Page_Portrait_Click()
        '<EhHeader>
        On Error GoTo Page_Portrait_Click_Err
        '</EhHeader>
100     Me.OASISDynamReports.SetOrientPortrait
102     Me.Page_Portrait.Checked = True
104     Me.Page_Landscape.Checked = False
        '<EhFooter>
        Exit Sub

Page_Portrait_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicReports.Page_Portrait_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

