VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Begin VB.Form frmIncidentVictimAnalysis 
   Caption         =   "Incident Victim Analysis"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9765
   Icon            =   "frmIncidentVictimAnalysis.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9765
      _cx             =   17224
      _cy             =   10081
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
      GridRows        =   3
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmIncidentVictimAnalysis.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame1 
         Height          =   630
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   9585
         Begin VB.CommandButton cmdExport 
            Caption         =   "Export to XLS"
            Height          =   255
            Left            =   7320
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdGenerate 
            Caption         =   "Generate"
            Height          =   255
            Left            =   5520
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
         Begin XpressEditorsLibCtl.dxDateEdit dxDateFrom 
            Height          =   315
            Left            =   840
            OleObjectBlob   =   "frmIncidentVictimAnalysis.frx":689D
            TabIndex        =   2
            Top             =   240
            Width           =   1815
         End
         Begin XpressEditorsLibCtl.dxDateEdit dxDateTill 
            Height          =   315
            Left            =   3480
            OleObjectBlob   =   "frmIncidentVictimAnalysis.frx":6973
            TabIndex        =   3
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblFrom 
            Caption         =   "From:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblTill 
            Caption         =   "Till:"
            Height          =   255
            Left            =   2880
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   4845
         Left            =   90
         OleObjectBlob   =   "frmIncidentVictimAnalysis.frx":6A49
         TabIndex        =   8
         Top             =   780
         Width           =   9585
      End
   End
End
Attribute VB_Name = "frmIncidentVictimAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSVictimData As ADODB.Recordset
Dim oConn As ADODB.Connection


Private Sub cmdExport_Click()
    Dim c As New cCommonDialog

    With c
        .DialogTitle = "Export to location..."
        .CancelError = False
        .hwnd = Me.hwnd
        .Flags = OFN_PATHMUSTEXIST
        .InitDir = g_sAppPath
        .Filter = "Excel File|*.xls"
        .FilterIndex = 1
        .ShowSave
        
    End With
        
    If Not IsNull(c.Filename) Then Me.dxDBGrid1.m.ExportToXLS IIf(Not (Right(c.Filename, 4)) = ".xls", c.Filename & ".xls", c.Filename)

End Sub

Private Sub cmdGenerate_Click()

    Dim sSQL As String
    Set ooConn = New ADODB.Connection
    Set RSVictimData = New ADODB.Recordset
    
    sSQL = "SELECT oincidents_FEA.Province, oincidents_FEA.District, oincidents_FEA.Town, " & "oincidents_FEA.TYPE AS IncidentType, oincidents_FEA.TARGET AS IncidentTarget, " & "oincidents_FEA.Violent AS IncidentViolent, IncidentVictims.Occupation AS VictimOccupation, " & "IncidentVictims.Under18 AS VictimUnder18, IncidentVictims.Sex AS VictimSex, " & "IncidentVictims.Condition AS VictimCondition, IncidentVictims.Ethnicity AS VictimEthnicity, IncidentVictims.Quantity AS Quantity " & "FROM IncidentVictims LEFT JOIN oincidents_FEA ON IncidentVictims.incidentID = oincidents_FEA.ID " & "WHERE (((oincidents_FEA.Incident_DATE)>=#" & Me.dxDateFrom & "# AND (oincidents_FEA.Incident_DATE)<=#" & Me.dxDateTill & "#));"

    With ooConn
        .CursorLocation = adUseClient
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & getClientDBPath & "\Oasisclient.mdb"
        .Open
    End With
        
    With RSVictimData
        Set .ActiveConnection = ooConn
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Source = sSQL
        .CursorLocation = adUseClient
        .Open
    End With
    
    Me.dxDBGrid1.Columns.DestroyColumns
    Set Me.dxDBGrid1.DataSource = RSVictimData
    Me.dxDBGrid1.Columns.RetrieveFields
    
End Sub

Private Function getClientDBPath()
        '<EhHeader>
        On Error GoTo getClientDBPath_Err
        '</EhHeader>

        Dim ofs As New FileSystemObject

100     If ofs.FolderExists(SpecialFolderPath(sftUserApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
102         getClientDBPath = SpecialFolderPath(sftUserApplicationData) & "iMMAP - OASIS\OASIS client\Data\db"
104     ElseIf ofs.FolderExists(SpecialFolderPath(sftUserLocalApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
106         getClientDBPath = SpecialFolderPath(sftUserLocalApplicationData) & "iMMAP - OASIS\OASIS client\db"
108     ElseIf ofs.FolderExists(SpecialFolderPath(sftUserMyDocuments) & "iMMAP - OASIS\OASIS client\Data\DB") Then
110         getClientDBPath = SpecialFolderPath(sftUserMyDocuments) & "iMMAP - OASIS\OASIS client\Data\db"
112     ElseIf ofs.FolderExists(SpecialFolderPath(sftCommonApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
114         getClientDBPath = SpecialFolderPath(sftCommonApplicationData) & "iMMAP - OASIS\OASIS client\Data\db"
116     ElseIf ofs.FolderExists(SpecialFolderPath(sftCommonMyDocuments) & "iMMAP - OASIS\OASIS client\Data\DB") Then
118         getClientDBPath = SpecialFolderPath(sftCommonMyDocuments) & "iMMAP - OASIS\OASIS client\Data\db"
        End If
    
        '<EhFooter>
        Exit Function

getClientDBPath_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.getClientDBPath " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub Form_Load()

    Dim dDate As Date
    
    dDate = Format(Now() - 365, "dd-MM-yy")
    Me.dxDateFrom = dDate
    
    dDate = Format(Now(), "dd-MM-yy")
    Me.dxDateTill = dDate
    
    If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
    
End Sub
