VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Begin VB.Form frmIncidentsMaintainance 
   Caption         =   "Incident Maintainance"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9690
   Icon            =   "frmIncidentsMaintainance.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5010
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   5010
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9690
      _cx             =   17092
      _cy             =   8837
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
      GridRows        =   4
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmIncidentsMaintainance.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridIncidents 
         Height          =   3420
         Left            =   90
         OleObjectBlob   =   "frmIncidentsMaintainance.frx":68B6
         TabIndex        =   1
         Top             =   1140
         Width           =   9510
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   300
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4620
         Width           =   9510
         _cx             =   16775
         _cy             =   529
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
         AutoSizeChildren=   7
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
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   300
            Left            =   8160
            TabIndex        =   4
            Top             =   0
            Width           =   1320
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   300
            Left            =   6675
            TabIndex        =   3
            Top             =   0
            Width           =   1350
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   465
         Left            =   90
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   615
         Width           =   9510
         _cx             =   16775
         _cy             =   820
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
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   2
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
         Begin VB.CommandButton cmdLoadByDate 
            Caption         =   "Load by Date"
            Height          =   285
            Left            =   6885
            TabIndex        =   6
            Top             =   90
            Width           =   2535
         End
         Begin XpressEditorsLibCtl.dxDateEdit dateFrom 
            Height          =   315
            Left            =   3105
            OleObjectBlob   =   "frmIncidentsMaintainance.frx":755E
            TabIndex        =   7
            Top             =   90
            Width           =   1710
         End
         Begin XpressEditorsLibCtl.dxDateEdit dateTill 
            Height          =   315
            Left            =   5265
            OleObjectBlob   =   "frmIncidentsMaintainance.frx":75FE
            TabIndex        =   8
            Top             =   90
            Width           =   1560
         End
         Begin VB.Label lblAnd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "and"
            Height          =   285
            Left            =   4875
            TabIndex        =   10
            Top             =   90
            Width           =   330
         End
         Begin VB.Label lblShowIncidents 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show incidents which occurred between"
            Height          =   285
            Left            =   90
            TabIndex        =   9
            Top             =   90
            Width           =   2955
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   465
         Left            =   90
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   90
         Width           =   9510
         _cx             =   16775
         _cy             =   820
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
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   2
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
         Begin VB.TextBox txtSQL 
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Text            =   "[Name] = 'TEST' OR [ID] LIKE '{0D4BA729%'"
            Top             =   90
            Width           =   6735
         End
         Begin VB.CommandButton cmdLoadBySQL 
            Caption         =   "Load by SQL Query"
            Height          =   285
            Left            =   6885
            TabIndex        =   12
            Top             =   90
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "frmIncidentsMaintainance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSIncidents As ADODB.Recordset
Dim bIncidentsLoaded As Boolean

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmIncidentsMaintainance.cmdCancel_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadData(sCondition As String)
        '<EhHeader>
        On Error GoTo LoadData_Err
        '</EhHeader>

        Dim cCol As dxGridColumn
100     C1Elastic1.Enabled = False
        Dim sString As String
        Dim oRS As ADODB.Recordset
        Dim i As Long
    
102     bIncidentsLoaded = False
104     Set dxDBGridIncidents.DataSource = Nothing
106     dxDBGridIncidents.Columns.DestroyColumns
108     cmdLoadBySQL.Enabled = False
110     cmdLoadByDate.Enabled = False

112     sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM oincidents " & sCondition)
114     Set oRS = m_frmOASISProgress.OpenHttpCommsRS(sString, True)
116     Set RSIncidents = CloneRS(oRS)
        RSIncidents.fields.Append "DeleteData", adBoolean
118     RSIncidents.open

120     With RSIncidents
        
122         Do Until oRS.EOF
124             RSIncidents.AddNew

                i = 0
                Do Until i >= oRS.fields.Count
                    If Not IsNull(oRS.fields(i).Value) Then .fields(oRS.fields(i).Name).Value = oRS.fields(i).Value
                    i = i + 1
                Loop
                oRS.MoveNext
            Loop
        
        End With
        
146     If Not RSIncidents.EOF Or Not RSIncidents.Bof Then

150         RSIncidents.MoveFirst

152         Set dxDBGridIncidents.DataSource = RSIncidents
154         dxDBGridIncidents.Columns.RetrieveFields
156         dxDBGridIncidents.Columns(0).Visible = False
158         dxDBGridIncidents.Columns(1).Visible = False
160         dxDBGridIncidents.Columns(2).Visible = False
162         dxDBGridIncidents.Columns(3).Visible = False
164         dxDBGridIncidents.Columns(4).Visible = False
166         dxDBGridIncidents.Columns(5).ReadOnly = True
168         bIncidentsLoaded = True
180     dxDBGridIncidents.KeyField = "ID"
        dxDBGridIncidents.Columns.ColumnByFieldName("DeleteData").Width = 3500
        
        End If
        

182     cmdLoadBySQL.Enabled = True
184     cmdLoadByDate.Enabled = True

186     C1Elastic1.Enabled = True

        '<EhFooter>
        Exit Sub

LoadData_Err:
        MsgBox "Error loading... please check your connection"
        cmdLoadBySQL.Enabled = True
        cmdLoadByDate.Enabled = True
        
        C1Elastic1.Enabled = True
        '</EhFooter>
End Sub

Private Function CloneRS(oRS As ADODB.Recordset, Optional sExcludedField As String = "") As ADODB.Recordset
        '<EhHeader>
        On Error GoTo CloneRS_Err
        '</EhHeader>

        Dim oRSClone As New ADODB.Recordset
        Dim i As Integer
100     i = 0
    
102     Do Until i = oRS.fields.Count
    
104         With oRS.fields(i)
        
106             If Not .Name = sExcludedField Then oRSClone.fields.Append .Name, .Type, .DefinedSize

            End With

108         i = i + 1
        Loop
    
110     Set CloneRS = oRSClone

        '<EhFooter>
        Exit Function

CloneRS_Err:
        MsgBox "CloneRS_Err on line (" & Erl & ") " & Err.Description
        
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdLoadByDate_Click()
    LoadData "WHERE Incident_DATESERIAL >= " & ConvertDateToSerial(dateFrom) & " AND  Incident_DATESERIAL <= " & ConvertDateToSerial(dateTill)

End Sub

Private Sub cmdLoadBySQL_Click()
    LoadData "WHERE " & txtSQL
End Sub

Private Sub cmdSave_Click()

    C1Elastic1.Enabled = False

    Dim bReturnValue As Boolean
    Dim iEdits As Integer
    Dim iDeletions As Integer
    Dim sDone As String
    Dim sSQL As String
        
    If RSIncidents Is Nothing Then Exit Sub
    Set dxDBGridIncidents.DataSource = Nothing
    dxDBGridIncidents.Visible = False
    Me.Refresh
    RSIncidents.Filter = "DeleteData = TRUE"
    bIncidentsLoaded = False
    
    If Not RSIncidents.EOF Or Not RSIncidents.Bof Then
    
        If MsgBox("Are you sure you want to delete " & RSIncidents.RecordCount & " records?", vbYesNo, "Confirm Deletion") = vbYes Then
            
            RSIncidents.MoveFirst

            Do Until RSIncidents.EOF

                sSQL = WebSite & "Oasis.asp?flagDeleteInSynchHist=" & RSIncidents.fields("ID").Value
                sDone = m_frmOASISProgress.OpenHttpCommsResponse(sSQL, True)

                RSIncidents.MoveNext
            Loop
                
            Me.Hide
            Unload Me

        End If
    End If
        
    dxDBGridIncidents.Visible = True
    Set dxDBGridIncidents.DataSource = Nothing
    
End Sub


Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     bIncidentsLoaded = False
102     Me.dateFrom = Format(Now() - 7, "Medium Date")
104     Me.dateTill = Format(Now(), "Medium Date")
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmIncidentsMaintainance.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
102     Set RSIncidents = Nothing
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmIncidentsMaintainance.Form_Unload " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


