VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{17BB99B1-05ED-474A-886C-DAD036A860F5}#3.0#0"; "OASISDynamicRSEditor.ocx"
Begin VB.Form frmDynamicData 
   Caption         =   "OASIS Dynamic Data Entry"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9480
   Icon            =   "frmDynamicData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9480
      _cx             =   16722
      _cy             =   11668
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
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   4
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
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmDynamicData.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   3645
         Left            =   4260
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   345
         Width           =   5160
         _cx             =   9102
         _cy             =   6429
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
         AutoSizeChildren=   7
         BorderWidth     =   3
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
         Begin OASISDynamRSEditor.OASISDynamicRSEditor OASISDynamicRSEditor1 
            Height          =   3555
            Index           =   0
            Left            =   45
            TabIndex        =   10
            Top             =   45
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   6271
         End
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H80000014&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4260
         MaskColor       =   &H0080FFFF&
         TabIndex        =   7
         Top             =   4050
         Visible         =   0   'False
         Width           =   2550
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   2160
         Left            =   60
         OleObjectBlob   =   "frmDynamicData.frx":68D3
         TabIndex        =   6
         Top             =   4395
         Width           =   9360
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6870
         TabIndex        =   5
         Top             =   4050
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit Selected Record"
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   4050
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.CommandButton cmdAddNew 
         Appearance      =   0  'Flat
         Caption         =   "Add New Record"
         Height          =   285
         Left            =   60
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4050
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00808080&
         ForeColor       =   &H0080FFFF&
         Height          =   3570
         Left            =   60
         TabIndex        =   1
         Top             =   345
         Width           =   4140
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   2160
         Left            =   60
         TabIndex        =   11
         Top             =   4395
         Width           =   9360
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0050C0A4&
         Caption         =   "Data Available"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   4140
      End
      Begin VB.Label lblOperation 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Browse records"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   4260
         TabIndex        =   4
         Top             =   60
         Width           =   5160
      End
   End
End
Attribute VB_Name = "frmDynamicData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type DYNAMIC_DATA_LIST_INFO
    Text As String
    DDDefIndex As Long
    DDDefTableIndex As Long
    FullTableName As String
    TablePrefix As String
    Excludes As String
    IsMasterTable As Boolean
End Type

Private Type DYNAMIC_DATA_DEF_TABLES
    TableName As String
    AllowRead As Boolean
    AllowAppend As Boolean
    AllowEdit As Boolean
    AllowDelete As Boolean
    Excludes As String
    IsMasterTable As Boolean
    IsLinkedTable As Boolean
End Type

Private Type DYNAMIC_DATA_DEF
    Name As String
    Desc As String
    Tables() As DYNAMIC_DATA_DEF_TABLES
End Type

Private mRS As ADODB.Recordset
Private mRS_GEO As ADODB.Recordset
Private mConn As ADODB.Connection
Private iCountOfGUIDs As Integer
Private sFilterBeforeAdd As String
Private sGUID1Current As String
Private sGUID2Current As String
Private sGUID1Old As String
Private sGUID2Old As String
Private sTableNames() As String
Private sLinkedTableNames() As String
Private sMasterTableName As String
Private bGeoTableActive As Boolean

Private ListInfo() As DYNAMIC_DATA_LIST_INFO
Private ListInfoCurrent As DYNAMIC_DATA_LIST_INFO
Private DynamicDataDefs() As DYNAMIC_DATA_DEF
Private DynamicDataDefCurrent As DYNAMIC_DATA_DEF
Private DynamicDataTableCurrent As DYNAMIC_DATA_DEF_TABLES

Private Sub DisplayGridData()
        '<EhHeader>
        On Error GoTo DisplayGridData_Err
        '</EhHeader>

        Dim lTimeToComplete As Long

100     If mRS.RecordCount > 1 Then
102         lTimeToComplete = ((48) / (31530)) * mRS.RecordCount
104         lblProgress.Caption = Chr(13) & "please wait approximately " & lTimeToComplete & " seconds for the data to load............"
106         lblProgress.Refresh
        End If
    
108     dxDBGrid1.Visible = False
110     List1.Enabled = False
112     cmdEdit.Visible = False
114     cmdAddNew.Visible = False
116     C1Elastic1.Refresh
    
118     Set dxDBGrid1.DataSource = Nothing
120     dxDBGrid1.Columns.DestroyColumns

122     If Not mRS.EOF And Not mRS.BOF Then mRS.Filter = mRS.Fields(0).Name & " = '" & mRS.Fields(0).Value & "'"
124     Set dxDBGrid1.DataSource = mRS

126     dxDBGrid1.Columns.RetrieveFields
128     dxDBGrid1.KeyField = mRS.Fields(0).Name
130     ScanForDDFields
132     Set dxDBGrid1.DataSource = Nothing
134     Set dxDBGrid1.DataSource = mRS
136     mRS.Filter = adFilterNone

138     If Not mRS.EOF Or Not mRS.BOF Then mRS.MoveFirst
140     If iCountOfGUIDs = 2 Then dxDBGrid1.Columns(1).Width = 50
142     If iCountOfGUIDs = 1 Or iCountOfGUIDs = 2 Then dxDBGrid1.Columns(0).Width = 50

144     dxDBGrid1.Visible = True
146     C1Elastic1.Refresh
148     List1.Enabled = True
150     Call SetAccessRights
    
152     lblProgress.Caption = ""
    
        '<EhFooter>
        Exit Sub

DisplayGridData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.DisplayGridData " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAddNew_Click()
        '<EhHeader>
        On Error GoTo cmdAddNew_Click_Err
        '</EhHeader>
        
        Dim sGUID As String
        Dim lUIDMax As Long

100     List1.Enabled = False
102     sFilterBeforeAdd = mRS.Filter
104     mRS.Filter = ""
106     dxDBGrid1.Visible = False
108     Call SetAccessRights
110     cmdCancel.Visible = True
112     iCountOfGUIDs = 1
114     lblOperation.BackColor = vbRed
116     lblOperation.Caption = "Current operation: ADD"
118     mRS.AddNew
120     C1Elastic2.BackColor = vbRed
122     OASISDynamicRSEditor1(0).SetGUID1 GUIDGen

124     If bGeoTableActive Then
126         mRS_GEO.AddNew
128         lUIDMax = GetMaxUID(ListInfoCurrent.FullTableName) + 1
130         OASISDynamicRSEditor1(0).SetUID (lUIDMax + 1)
132         OASISDynamicRSEditor1(1).SetUID (lUIDMax + 1)
        End If

134     OASISDynamicRSEditor1(0).LockUnlockAllFields False
    
        '<EhFooter>
        Exit Sub

cmdAddNew_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.cmdAddNew_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetMaxUID(sTableName As String) As Long
        '<EhHeader>
        On Error GoTo GetMaxUID_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
100     oRS.Open "SELECT MAX(UID) AS DUDE FROM [" & sTableName & "]", mConn, adOpenDynamic, adLockBatchOptimistic
102     GetMaxUID = CLng(oRS.Fields(0).Value)
104     oRS.Close
106     Set oRS = Nothing
    
        '<EhFooter>
        Exit Function

GetMaxUID_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.GetMaxUID " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
        
        Dim bSaveIsVisible As Boolean

100     cmdCancel.Visible = False
102     dxDBGrid1.Enabled = True
104     Me.SetFocus
106     bSaveIsVisible = cmdSave.Visible
108     cmdSave.Visible = False
110     C1Elastic1.Refresh
    
112     If bSaveIsVisible And dxDBGrid1.Visible Then
            'edit took place and needs to be reverted
114         mRS.Requery

116         If bGeoTableActive Then mRS_GEO.Requery
        End If

118     cmdSave.Visible = False
120     Me.SetFocus

122     If Not dxDBGrid1.Visible Then
            'Add cancelled
124         dxDBGrid1.Visible = True
126         mRS.Requery

128         If bGeoTableActive Then mRS_GEO.Requery
130         If Not sFilterBeforeAdd = "0" Then mRS.Filter = sFilterBeforeAdd
        End If

132     lblOperation.BackColor = &HFF8080
134     lblOperation.Caption = "Browse records"
136     C1Elastic2.BackColor = -2147483633
138     OASISDynamicRSEditor1(0).LockUnlockAllFields True
    
140     Call SetAccessRights
142     cmdSave.Visible = False
144     List1.Enabled = True

        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSave_Click()
        '<EhHeader>
        On Error GoTo cmdSave_Click_Err
        '</EhHeader>

100     If OASISDynamicRSEditor1(0).bChangeMade Then
        
102         If dxDBGrid1.Visible Then
            
104             If iCountOfGUIDs = 2 Then
                
106                 sGUID2Old = mRS.Fields(1).Value
108                 sGUID2Current = GUIDGen
110                 mRS.Fields(1).Value = sGUID2Current
                    
                End If
                
112             sGUID1Old = mRS.Fields(0).Value
114             sGUID1Current = GUIDGen
116             mRS.Fields(0).Value = sGUID1Current
            
            End If
        
118         OASISDynamicRSEditor1(0).SaveRecord
120         MsgBox "Saved"

        Else
122         MsgBox "Nothing to save"
        End If
    
124     dxDBGrid1.Enabled = True
126     Me.SetFocus
128     cmdSave.Visible = False
130     Call SetAccessRights
132     cmdCancel.Visible = False
    
134     If Not dxDBGrid1.Visible Then
136         dxDBGrid1.Visible = True

138         If Not sFilterBeforeAdd = "0" Then mRS.Filter = sFilterBeforeAdd
        End If
    
140     lblOperation.BackColor = &HFF8080
142     lblOperation.Caption = "Browse records"
144     C1Elastic2.BackColor = -2147483633
     
146     OASISDynamicRSEditor1(0).LockUnlockAllFields True
148     List1.Enabled = True
    
        '<EhFooter>
        Exit Sub

cmdSave_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.cmdSave_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdEdit_Click()
        '<EhHeader>
        On Error GoTo cmdEdit_Click_Err
        '</EhHeader>

100     If mRS.EOF Or mRS.BOF Then
102         MsgBox "You must select a record before you can edit", vbExclamation, "Edit record"
        Else
104         List1.Enabled = False
106         dxDBGrid1.Enabled = False
108         Call SetAccessRights
110         cmdCancel.Visible = True
112         iCountOfGUIDs = 1
114         lblOperation.BackColor = vbRed
116         lblOperation.Caption = "Current operation: EDIT"
118         C1Elastic2.BackColor = vbRed

120         If bGeoTableActive Then
            
122             SetGeoRsToCurrentUID
            
            End If

124         OASISDynamicRSEditor1(0).LockUnlockAllFields False
        End If

        '<EhFooter>
        Exit Sub

cmdEdit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.cmdEdit_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetGeoRsToCurrentUID()
        '<EhHeader>
        On Error GoTo SetGeoRsToCurrentUID_Err
        '</EhHeader>
    
100     If mRS_GEO.EOF Or mRS_GEO.BOF Then
    
102         If Not mRS_GEO.EOF Or Not mRS_GEO.BOF Then
104             mRS_GEO.MoveFirst
106             mRS_GEO.Find ("UID = " & mRS.Fields("UID").Value)
            End If
    
108     ElseIf Not mRS_GEO.Fields("UID").Value = mRS.Fields("UID").Value Then
    
110         If Not mRS_GEO.EOF Or Not mRS_GEO.BOF Then
112             mRS_GEO.MoveFirst
114             mRS_GEO.Find ("UID = " & mRS.Fields("UID").Value)
            End If
        
        End If

        '<EhFooter>
        Exit Sub

SetGeoRsToCurrentUID_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.SetGeoRsToCurrentUID " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDBGrid1_OnDblClick()
        '<EhHeader>
        On Error GoTo dxDBGrid1_OnDblClick_Err
        '</EhHeader>

100     If DynamicDataTableCurrent.AllowDelete Then

102         If Not mRS.EOF And Not mRS.BOF Then
    
104             If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Confirm deletion") = vbYes Then
106                 mRS.Filter = mRS.Fields(0).Name & " = '" & mRS.Fields(0).Value & "'"
108                 dxDBGrid1.Dataset.Edit
110                 dxDBGrid1.Dataset.Delete
112                 mRS.UpdateBatch
114                 mRS.Filter = adFilterNone
                End If
    
            End If
    
        End If

        '<EhFooter>
        Exit Sub

dxDBGrid1_OnDblClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.dxDBGrid1_OnDblClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub DisplayGEOTable(oRS As ADODB.Recordset, _
                            sTableNamePrefix As String)
        '<EhHeader>
        On Error GoTo DisplayGEOTable_Err
        '</EhHeader>

100     OASISDynamicRSEditor1(0).Height = 2 * (OASISDynamicRSEditor1(0).Height / 3)
102     Load OASISDynamicRSEditor1(1)
104     OASISDynamicRSEditor1(1).Top = OASISDynamicRSEditor1(0).Top + OASISDynamicRSEditor1(0).Height
106     OASISDynamicRSEditor1(1).Height = OASISDynamicRSEditor1(0).Height / 2
108     OASISDynamicRSEditor1(1).Visible = True
110     SetGeoRsToCurrentUID
112     OASISDynamicRSEditor1(1).Init oRS, mConn, sTableNamePrefix, vbGrayText, vbWhite, , True
114     OASISDynamicRSEditor1(1).LockUnlockAllFields True
    
        '<EhFooter>
        Exit Sub

DisplayGEOTable_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.DisplayGEOTable " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub RemoveGEOTable()
        '<EhHeader>
        On Error GoTo RemoveGEOTable_Err
        '</EhHeader>

100     If OASISDynamicRSEditor1.Count = 2 Then
    
102         OASISDynamicRSEditor1(0).Height = 3 * (OASISDynamicRSEditor1(0).Height / 2)
104         Unload OASISDynamicRSEditor1(1)
    
        End If
    
        '<EhFooter>
        Exit Sub

RemoveGEOTable_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.RemoveGEOTable " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        
        Dim I As Long
        Dim l As Long
        Dim iDDDefIndex As Long

        Dim oRS As New ADODB.Recordset
100     Set mConn = New ADODB.Connection
102     Set mRS = New ADODB.Recordset
    
104     sMasterTableName = "mastertable"
    
106     mConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\IMMAP\Documents\iMMAP - OASIS\OASIS client\data\db\OasisClient.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False;"
108     mConn.CursorLocation = adUseClient
110     mConn.Open

112     oRS.Open "SELECT * FROM DynamicDataDefs", mConn, adOpenDynamic, adLockReadOnly
    
114     ReDim DynamicDataDefs(0)
116     ReDim ListInfo(0)
    
118     l = 0
120     iDDDefIndex = 0
    
122     Do Until oRS.EOF
    
124         iDDDefIndex = iDDDefIndex + 1
        
126         ReDim Preserve DynamicDataDefs(UBound(DynamicDataDefs) + 1)
128         DynamicDataDefs(UBound(DynamicDataDefs)).Name = oRS.Fields("DDDefName").Value
130         DynamicDataDefs(UBound(DynamicDataDefs)).Desc = oRS.Fields("Description").Value
            
132         GetDDDetailedInfo "dd_" & oRS.Fields("DDDefName").Value & "_", IIf(IsNull(oRS.Fields("ExcludedFields").Value), "", oRS.Fields("ExcludedFields").Value), IIf(IsNull(oRS.Fields("AccessRights").Value), "", oRS.Fields("AccessRights").Value), DynamicDataDefs(UBound(DynamicDataDefs))
           
134         I = 1

136         Do Until I = UBound(DynamicDataDefs(UBound(DynamicDataDefs)).Tables) + 1
                
138             If Not Left$(DynamicDataDefs(UBound(DynamicDataDefs)).Tables(I).TableName, 4) = "link" And DynamicDataDefs(UBound(DynamicDataDefs)).Tables(I).AllowRead And Not Right$(DynamicDataDefs(UBound(DynamicDataDefs)).Tables(I).TableName, 4) = "_GEO" Then
                
140                 If DynamicDataDefs(UBound(DynamicDataDefs)).Tables(I).AllowRead Then
                
142                     ReDim Preserve ListInfo(UBound(ListInfo) + 1)

144                     With ListInfo(UBound(ListInfo))
                
146                         .Text = DynamicDataDefs(UBound(DynamicDataDefs)).Desc & " (" & Right$(DynamicDataDefs(UBound(DynamicDataDefs)).Tables(I).TableName, Len(DynamicDataDefs(UBound(DynamicDataDefs)).Tables(I).TableName) - 2) & ")"
                
148                         If DynamicDataDefs(UBound(DynamicDataDefs)).Tables(I).TableName = "mastertable" Then
150                             .Text = DynamicDataDefs(UBound(DynamicDataDefs)).Desc & " (DATA ENTRY)"
                            End If
                
152                         .FullTableName = "dd_" & DynamicDataDefs(UBound(DynamicDataDefs)).Name & "_" & DynamicDataDefs(UBound(DynamicDataDefs)).Tables(I).TableName
154                         .DDDefIndex = iDDDefIndex
156                         .DDDefTableIndex = I
158                         .TablePrefix = "dd_" & DynamicDataDefs(UBound(DynamicDataDefs)).Name & "_"
160                         .Excludes = DynamicDataDefs(UBound(DynamicDataDefs)).Tables(I).Excludes
162                         .IsMasterTable = DynamicDataDefs(UBound(DynamicDataDefs)).Tables(I).IsMasterTable
164                         List1.AddItem .Text
166                         l = l + 1
                    
                        End With
                
                    End If

                End If
            
168             I = I + 1
           
            Loop
    
170         oRS.MoveNext
        
        Loop

172     OASISDynamicRSEditor1(0).Reset
174     Me.WindowState = 2

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
      
        On Error Resume Next
100     Set mRS = Nothing
102     Set mConn = Nothing

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub OASISDynamicRSEditor1_SoundChangeMade(Index As Integer)
        '<EhHeader>
        On Error GoTo OASISDynamicRSEditor1_SoundChangeMade_Err
        '</EhHeader>

100     If lblOperation.BackColor = vbRed Then
    
102         If OASISDynamicRSEditor1(0).bChangeMade Then
104             cmdSave.Visible = True
            Else
106             cmdSave.Visible = False
            End If
        
        End If
        
108     If bGeoTableActive Then Call SetGeoRsToCurrentUID
    
        '<EhFooter>
        Exit Sub

OASISDynamicRSEditor1_SoundChangeMade_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.OASISDynamicRSEditor1_SoundChangeMade " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ScanForDDFields()
        '<EhHeader>
        On Error GoTo ScanForDDFields_Err
        '</EhHeader>

        Dim I As Long
100     I = 0
    
102     Do Until I = mRS.Fields.Count
    
104         If Left(mRS.Fields(I).Name, 2) = "dd" Then
106             SetDropdown I
            End If

108         I = I + 1
        Loop

        '<EhFooter>
        Exit Sub

ScanForDDFields_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.ScanForDDFields " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetDropdown(iColumnIndex As Long)
        '<EhHeader>
        On Error GoTo SetDropdown_Err
        '</EhHeader>

100     dxDBGrid1.Columns(iColumnIndex).ColumnType = gedLookupEdit
    
102     With dxDBGrid1.Columns(iColumnIndex).LookupColumn
        
104         .LookupDatasetType = dtADODataset

106         With .LookupDataset.ADODataset
108             .ConnectionString = mConn.ConnectionString '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Documents\MCDB_backend.mdb;Persist Security Info=False"
110             .CursorLocation = clUseServer
112             .CursorType = ctKeyset  'ctStatic
114             .CommandType = cmdTable
116             .CommandText = "[" & ListInfoCurrent.TablePrefix & mRS.Fields(iColumnIndex).Name & "]"   'dxDBGrid1.Columns(iColumnIndex).FieldName
118             .LockType = ltBatchOptimistic
            End With

120         .LookupDataset.Open
122         .DisplaySize = 20
124         .ListAutoWidth = False
126         .ListColumns = "*"
128         .ListFieldIndex = 0
130         .ListFieldName = "option"
132         .ListWidth = 0
134         .LookupCache = True
136         .LookupDataset.Active = False
138         .LookupDatasetType = dtADODataset
140         .LookupKeyField = "id"
142         .LookupResultField = "option"
144         .ListFieldIndex = 0
146         .ListWidth = 200
148         .ListColumns = ""

        End With
    
150     dxDBGrid1.Columns(iColumnIndex).Caption = GetFieldCaption(mConn, mRS, mRS.Fields(iColumnIndex).Name)

        '<EhFooter>
        Exit Sub

SetDropdown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.SetDropdown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetFieldCaption(oConn As ADODB.Connection, _
                                 RSLocalRecordset As ADODB.Recordset, _
                                 sFieldName As String)
        '<EhHeader>
        On Error GoTo GetFieldCaption_Err
        '</EhHeader>
 
        Dim oDB As ADOx.Catalog
        Dim itbl As ADOx.Table
        Dim fld As ADOx.Column
 
100     Set oDB = New ADOx.Catalog
102     Set itbl = New ADOx.Table
104     Set oDB.ActiveConnection = oConn
106     GetFieldCaption = "desc not defined"

108     For Each itbl In oDB.Tables

110         If itbl.Name = RSLocalRecordset.Fields(0).Properties(1) Then
112             GetFieldCaption = itbl.Columns(sFieldName).Properties(2).Value
                Exit For
            End If

        Next
        
114     Set oDB = Nothing
116     Set itbl = Nothing
118     Set fld = Nothing

        '<EhFooter>
        Exit Function

GetFieldCaption_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.GetFieldCaption " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub SetAccessRights()
        '<EhHeader>
        On Error GoTo SetAccessRights_Err
        '</EhHeader>

100     cmdAddNew.Visible = DynamicDataTableCurrent.AllowAppend
102     cmdEdit.Visible = DynamicDataTableCurrent.AllowEdit

        '<EhFooter>
        Exit Sub

SetAccessRights_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.SetAccessRights " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub List1_Click()
        '<EhHeader>
        On Error GoTo List1_Click_Err
        '</EhHeader>

        Dim sSQL As String
100     bGeoTableActive = False
102     Call RemoveGEOTable
104     ListInfoCurrent = ListInfo(List1.ListIndex + 1)
        
106     DynamicDataTableCurrent = DynamicDataDefs(ListInfoCurrent.DDDefIndex).Tables(ListInfoCurrent.DDDefTableIndex)
108     iCountOfGUIDs = 1

110     If Not IsNull(List1.Text) And Not List1.Text = "" Then

112         sSQL = "SELECT * FROM [" & ListInfoCurrent.FullTableName & "]"
114         C1Elastic1.Refresh
116         Set mRS = New ADODB.Recordset
118         mRS.Open sSQL, mConn, adOpenDynamic, adLockBatchOptimistic
120         OASISDynamicRSEditor1(0).Init mRS, mConn, ListInfoCurrent.TablePrefix, vbGrayText, vbWhite, iCountOfGUIDs
122         OASISDynamicRSEditor1(0).LockUnlockAllFields True

124         If Right$(ListInfoCurrent.FullTableName, 4) = "_FEA" Then
126             bGeoTableActive = True
128             Set mRS_GEO = New ADODB.Recordset
130             sSQL = "SELECT * FROM [" & Left$(ListInfoCurrent.FullTableName, Len(ListInfoCurrent.FullTableName) - 4) & "_GEO]"
132             mRS_GEO.Open sSQL, mConn, adOpenDynamic, adLockBatchOptimistic
134             DisplayGEOTable mRS_GEO, ListInfoCurrent.TablePrefix
            End If
            
136         If bGeoTableActive Then SetGeoRsToCurrentUID
138         Call SetAccessRights
140         Call DisplayGridData
    
        End If

        '<EhFooter>
        Exit Sub

List1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.List1_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GetDDDetailedInfo(sTablePrefix As String, _
                              sExcludes As String, _
                              sAccessRights As String, _
                              DynamicDataDef As DYNAMIC_DATA_DEF)
        '<EhHeader>
        On Error GoTo GetDDDetailedInfo_Err
        '</EhHeader>

        Dim I As Long
        Dim l As Long
        Dim sTableName As String
        Dim sAccessRightsArray() As String
        Dim sExcludeArray() As String
        Dim sSmallSplit() As String
        Dim sExcludeString As String
    
100     ReDim Preserve DynamicDataDef.Tables(0)
    
102     sAccessRightsArray = Split(sAccessRights, ";", -1, vbTextCompare)
104     sExcludeArray = Split(sExcludes, ";", -1, vbTextCompare)
106     I = 0

108     Do Until I = UBound(sAccessRightsArray)
    
110         ReDim Preserve DynamicDataDef.Tables(UBound(DynamicDataDef.Tables) + 1)
        
112         With DynamicDataDef.Tables(UBound(DynamicDataDef.Tables))
        
114             sSmallSplit = Split(sAccessRightsArray(I), ",", -1, vbTextCompare)
            
116             .TableName = sSmallSplit(0)
118             .AllowAppend = False
120             .AllowDelete = False
122             .AllowEdit = False
124             .AllowRead = False

126             If InStr(1, sSmallSplit(1), "r", vbTextCompare) Then .AllowRead = True
128             If InStr(1, sSmallSplit(1), "e", vbTextCompare) Then .AllowEdit = True
130             If InStr(1, sSmallSplit(1), "a", vbTextCompare) Then .AllowAppend = True
132             If InStr(1, sSmallSplit(1), "d", vbTextCompare) Then .AllowDelete = True
            
134             l = 0
136             sExcludeString = ""
            
138             If Not sExcludes = "" Then

140                 Do Until l = UBound(sExcludeArray)
                
142                     sSmallSplit = Split(sExcludeArray(l), ",", -1, vbTextCompare)

144                     If sSmallSplit(0) = .TableName Then
146                         sExcludeString = sExcludeString & sSmallSplit(1) & ","
                        End If

148                     l = l + 1
                    Loop

                End If
            
150             If Len(sExcludeString) > 1 Then sExcludeString = Left$(sExcludeString, Len(sExcludeString) - 1)
152             .Excludes = sExcludeString
            
            End With
        
154         I = I + 1
        Loop

        '<EhFooter>
        Exit Sub

GetDDDetailedInfo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RunnableProject.frmDynamicData.GetDDDetailedInfo " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
