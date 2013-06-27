VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmDatabaseExplorer 
   Caption         =   "OASIS Server Database Explorer"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9585
   Icon            =   "frmDatabaseExplorer.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7170
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7170
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9585
      _cx             =   16907
      _cy             =   12647
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
      PicturePos      =   10
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmDatabaseExplorer.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   1380
         Left            =   2700
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   6885
         _cx             =   12144
         _cy             =   2434
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
         Align           =   0
         AutoSizeChildren=   0
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
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "This form provides read only access to all tables in the OASIS Server Database"
            ForeColor       =   &H8000000B&
            Height          =   285
            Left            =   420
            TabIndex        =   5
            Top             =   930
            Width           =   5895
         End
         Begin VB.Label Label1 
            BackColor       =   &H0050C0A4&
            Caption         =   "Database Explorer"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   555
            Left            =   390
            TabIndex        =   3
            Top             =   270
            Width           =   6435
         End
      End
      Begin VB.ListBox listTables 
         Height          =   5715
         Left            =   75
         Sorted          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Click on table to view detail"
         Top             =   1440
         Width           =   2565
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxdbGrid1 
         Height          =   5730
         Left            =   2700
         OleObjectBlob   =   "frmDatabaseExplorer.frx":68D4
         TabIndex        =   4
         Top             =   1440
         Width           =   6885
      End
   End
End
Attribute VB_Name = "frmDatabaseExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSTables As ADODB.Recordset
Dim RSActiveTable As ADODB.Recordset

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     C1Elastic1.Picture = g_PictureDialogSmall
    
        If Right(WebSite, 1) <> "/" Then
            WebSite = WebSite & "/"
        End If
    
102     Set RSTables = m_frmOASISProgress.OpenHttpCommsRS(WebSite & "oasis.asp?gettables", True)

104     If Not RSTables Is Nothing Then
106         RSTables.Filter = "TABLE_TYPE = 'TABLE'"

108         Do Until RSTables.EOF
    
110             listTables.AddItem RSTables.fields("TABLE_NAME").Value
    
112             RSTables.MoveNext
            Loop
            
114         Set dxDBGrid1.DataSource = RSActiveTable
        Else
116         MsgBox "Server connection failed"
118         Me.Hide
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDatabaseExplorer.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub listTables_Click()
        '<EhHeader>
        On Error GoTo listTables_Click_Err
        '</EhHeader>
    
100     Set RSActiveTable = m_frmOASISProgress.OpenHttpCommsRS(WebSite & "oasis.asp?id=" & CheckEncrypt(listTables.Text), True)
    
102     If Not RSActiveTable Is Nothing Then
104         dxDBGrid1.Columns.DestroyColumns
106         dxDBGrid1.KeyField = RSActiveTable.fields(0).Name
108         Set dxDBGrid1.DataSource = Nothing
110         Set dxDBGrid1.DataSource = RSActiveTable.clone
112         dxDBGrid1.Columns.RetrieveFields
            
        Else
114         MsgBox "Server connection failed"

        End If

        '<EhFooter>
        Exit Sub

listTables_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDatabaseExplorer.listTables_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
