VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDynamicDataMenu 
   BackColor       =   &H0050C0A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dynamic Data Configuration Wizard"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8865
   Icon            =   "frmDynamicDataMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6165
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8865
      _cx             =   15637
      _cy             =   10874
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
      GridRows        =   5
      GridCols        =   5
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmDynamicDataMenu.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridAccessRights 
         Height          =   2460
         Left            =   4515
         OleObjectBlob   =   "frmDynamicDataMenu.frx":68EC
         TabIndex        =   8
         Top             =   3300
         Width           =   4260
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add new Dynamic Database"
         Height          =   2460
         Left            =   6675
         TabIndex        =   7
         Top             =   435
         Width           =   2100
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   4515
         TabIndex        =   3
         Top             =   5820
         Width           =   2100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   6675
         TabIndex        =   2
         Top             =   5820
         Width           =   2100
      End
      Begin MSComctlLib.ListView listExludes 
         Height          =   2775
         Left            =   90
         TabIndex        =   1
         Top             =   3300
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   4895
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridDefinedDBs 
         Height          =   2460
         Left            =   90
         OleObjectBlob   =   "frmDynamicDataMenu.frx":7594
         TabIndex        =   9
         Top             =   435
         Width           =   6525
      End
      Begin CONTROLSLibCtl.dxLabel dxLabel1 
         Height          =   285
         Index           =   1
         Left            =   4515
         TabIndex        =   6
         Top             =   2955
         Width           =   4260
         _Version        =   0
         _cx             =   7514
         _cy             =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Access rights for this dynamic database"
         BackStyle       =   1
         BackColor       =   5292196
         ForeColor       =   0
         LabelStyle      =   0
         Label3dStyle    =   2
         Label3dOrientation=   4
         Label3dDepth    =   0
         PenWidth        =   1
         Angle           =   0
         ShadowColor     =   10526880
      End
      Begin CONTROLSLibCtl.dxLabel dxLabel3 
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   2955
         Width           =   4365
         _Version        =   0
         _cx             =   7699
         _cy             =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Untick fields you dont want to use"
         BackStyle       =   1
         BackColor       =   5292196
         ForeColor       =   0
         LabelStyle      =   0
         Label3dStyle    =   2
         Label3dOrientation=   4
         Label3dDepth    =   0
         PenWidth        =   1
         Angle           =   0
         ShadowColor     =   10526880
      End
      Begin CONTROLSLibCtl.dxLabel dxLabel2 
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   90
         Width           =   8685
         _Version        =   0
         _cx             =   15319
         _cy             =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "All Defined Dynamic Databases"
         BackStyle       =   1
         BackColor       =   12648447
         ForeColor       =   0
         LabelStyle      =   0
         Label3dStyle    =   2
         Label3dOrientation=   4
         Label3dDepth    =   0
         PenWidth        =   1
         Angle           =   0
         ShadowColor     =   10526880
      End
   End
End
Attribute VB_Name = "frmDynamicDataMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RSAccessRights As ADODB.Recordset
Private RSDefinedDynDatabases As ADODB.Recordset
'Private oConn As ADODB.Connection
Private RemoteConn As ADODB.Connection

'Private sAvailableTableNames As String
Private sWebsite As String
Private sUserGroupName As String
Private sUserGroupGUID As String

Public Sub setUserGroupsRS(sPassedUserGroupName As String, _
                           sPassedUserGroupGUID As String)
        '<EhHeader>
        On Error GoTo setUserGroupsRS_Err
        '</EhHeader>

100     sUserGroupName = sPassedUserGroupName
102     sUserGroupGUID = sPassedUserGroupGUID

        '<EhFooter>
        Exit Sub

setUserGroupsRS_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.setUserGroupsRS " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAdd_Click()

    '
    '100     If Not listAvailable = "" And Not IsNull(listAvailable) Then
    '
    '102         RSDefinedDynDatabases.AddNew
    '104         RSDefinedDynDatabases.fields("DynamDBDefGUID").Value = GUIDGen
    '106         RSDefinedDynDatabases.fields("GroupGUID").Value = sUserGroupGUID
    '108         RSDefinedDynDatabases.fields("DDDefName").Value = listAvailable
    '110         RSDefinedDynDatabases.fields("Description").Value = "edit this description"
    '112         LoadListDatas
    '        End If
    '

    frmDialogWithTwoFields.Caption = "New Dynamic Data Def"
    frmDialogWithTwoFields.lbl1.Caption = "Please specify the prefix for each Dynamic Data Table" & Chr(13) & "(with no spaces as shown below in the textbox)"
    frmDialogWithTwoFields.lbl2.Caption = "Please specify the connection string" & Chr(13) & "(please use a relative path as shown below in the textbox)"
    frmDialogWithTwoFields.Height = frmDialogWithTwoFields.Height + (frmDialogWithTwoFields.txt2.Height * 2)
    frmDialogWithTwoFields.txt2.Height = frmDialogWithTwoFields.txt2.Height * 3
    frmDialogWithTwoFields.txt1 = "OCHA-NFI-Database"
    frmDialogWithTwoFields.txt2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\data\db\dynamicdata\DATABASENAME.MDB;Mode=ReadWrite|Share Deny None;Persist Security Info=False;"
    frmDialogWithTwoFields.Show vbModal, Me
        
    If frmDialogWithTwoFields.bClickedOK Then
        
        If Not frmDialogWithTwoFields.sText1 = "" And Not frmDialogWithTwoFields.sText2 = "" Then
                
            RSDefinedDynDatabases.AddNew
            RSDefinedDynDatabases.fields("DynamDBDefGUID").Value = GUIDGen
            RSDefinedDynDatabases.fields("GroupGUID").Value = sUserGroupGUID
            RSDefinedDynDatabases.fields("DDDefName").Value = frmDialogWithTwoFields.sText1
            RSDefinedDynDatabases.fields("Description").Value = "edit this description"
            RSDefinedDynDatabases.fields("ConnectionString").Value = frmDialogWithTwoFields.sText2
            
        End If
        
    End If
        
End Sub

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.cmdCancel_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetListAccessRights(sDDDefName As String, _
                                sAccessRights As String)
        '<EhHeader>
        On Error GoTo SetListAccessRights_Err
        '</EhHeader>
    
        Dim iINx() As Integer
        Dim sLocalGUID As String
        Dim sDDTableNamePrefix As String
        Dim sAccessRight1 As String
        
        'Dim oConn As New ADODB.Connection
        Dim oRs As New ADODB.Recordset
100     Set RSAccessRights = New ADODB.Recordset

102     'With oConn
104         '.CursorLocation = adUseClient
106         '.ConnectionString = GetConnectionString(CreateAppPath & "\data\db\Oasisclient.mdb")
108         '.open
        'End With
        
        
110     Set oRs = RemoteConn.OpenSchema(adSchemaTables)
        '  oRS.Sort = "TABLENAME DESC"
  
112     RSAccessRights.fields.Append "Table Name", adVarChar, 50
114     RSAccessRights.fields.Append "Read", adBoolean
116     RSAccessRights.fields.Append "Append", adBoolean
118     RSAccessRights.fields.Append "Edit", adBoolean
120     RSAccessRights.fields.Append "Delete", adBoolean
122     RSAccessRights.open
    
124     sLocalGUID = GUIDGen
        
126     sDDTableNamePrefix = "dd_" & sDDDefName & "_"

128     While Not oRs.EOF
         
130         If (Left$(oRs!TABLE_NAME, Len(sDDTableNamePrefix) + 4) = sDDTableNamePrefix & "link") Then
            
132             sTableName = Right(oRs!TABLE_NAME, Len(oRs!TABLE_NAME) - Len(sDDTableNamePrefix))
            
134         ElseIf oRs!TABLE_NAME = sDDTableNamePrefix & "mastertable" Then
            
136             sTableName = "mastertable"
            
138         ElseIf (Left$(oRs!TABLE_NAME, Len(sDDTableNamePrefix) + 2) = sDDTableNamePrefix & "dd") Then
            
140             sTableName = Right(oRs!TABLE_NAME, Len(oRs!TABLE_NAME) - Len(sDDTableNamePrefix))
            
            Else
            
142             sTableName = ""
            
            End If
            
144         If Not sTableName = "" Then
            
146             sLocalGUID = GUIDGen
148             RSAccessRights.AddNew
150             RSAccessRights.fields("Table Name").Value = sTableName
152             RSAccessRights.fields("Read").Value = False
154             RSAccessRights.fields("Append").Value = False
156             RSAccessRights.fields("Edit").Value = False
158             RSAccessRights.fields("Delete").Value = False

160             sAccessRight1 = sAccessRights
            
162             If InStr(1, sAccessRight1, sTableName, vbTextCompare) <> 0 Then
            
164                 If InStr(1, sAccessRight1, sTableName, vbTextCompare) + Len(sTableName) = Len(sAccessRight1) Then
                
166                     sAccessRight1 = ""
                
                    Else
                
168                     sAccessRight1 = Mid$(sAccessRight1, InStr(1, sAccessRight1, sTableName, vbTextCompare), InStr(InStr(1, sAccessRight1, sTableName, vbTextCompare), sAccessRight1, ";", vbTextCompare) - InStr(1, sAccessRight1, sTableName, vbTextCompare))
170                     sAccessRight1 = Right$(sAccessRight1, Len(sAccessRight1) - Len(sTableName) - 1)
                    End If

172                 If InStr(1, sAccessRight1, "r", vbTextCompare) <> 0 Then
174                     RSAccessRights.fields("Read").Value = True
                    End If
            
176                 If InStr(1, sAccessRight1, "a", vbTextCompare) <> 0 Then
178                     RSAccessRights.fields("Append").Value = True
                    End If
            
180                 If InStr(1, sAccessRight1, "e", vbTextCompare) <> 0 Then
182                     RSAccessRights.fields("Edit").Value = True
                    End If
            
184                 If InStr(1, sAccessRight1, "d", vbTextCompare) <> 0 Then
186                     RSAccessRights.fields("Delete").Value = True
                    End If
            
                End If

            End If
            
188         oRs.MoveNext
            
        Wend
    
190     Set dxDBGridAccessRights.DataSource = RSAccessRights
192     dxDBGridAccessRights.Columns.RetrieveFields
194     dxDBGridAccessRights.Columns(0).ReadOnly = True
    
        '<EhFooter>
        Exit Sub

SetListAccessRights_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.SetListAccessRights " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetListExcludes(sDDDefName As String, _
                            sExcludedFields As String)
        '<EhHeader>
        On Error GoTo SetListExcludes_Err
        '</EhHeader>
    
        Dim iINx() As Integer
        Dim sLocalGUID As String
        Dim sDDTableNamePrefix As String
        
        'Dim oConn As New ADODB.Connection
        Dim oRs As New ADODB.Recordset
        Dim sTableNameOld As String
        Dim iCountOfLinkedPKs As Integer

100     'With oConn
102         '.CursorLocation = adUseClient
104         '.ConnectionString = GetConnectionString(CreateAppPath & "\data\db\Oasisclient.mdb")
106         '.open
        'End With
    
108     Set oRs = RemoteConn.OpenSchema(adSchemaColumns)
110     oRs.Sort = "TABLE_NAME, ORDINAL_POSITION"
        'oRs.Sort = ""

112     listExludes.ListItems.Clear

114     With Me.listExludes

116         sLocalGUID = GUIDGen
        
118         sDDTableNamePrefix = "dd_" & sDDDefName & "_"
        
120         While Not oRs.EOF
         
122             If (Left$(oRs!TABLE_NAME, Len(sDDTableNamePrefix) + 4) = sDDTableNamePrefix & "link") Then
            
124                 sTableName = Right(oRs!TABLE_NAME, Len(oRs!TABLE_NAME) - Len(sDDTableNamePrefix))
            
126             ElseIf oRs!TABLE_NAME = sDDTableNamePrefix & "mastertable" Then
            
128                 sTableName = "mastertable"
            
130             ElseIf (Left$(oRs!TABLE_NAME, Len(sDDTableNamePrefix) + 2) = sDDTableNamePrefix & "dd") Then
            
132                 sTableName = Right(oRs!TABLE_NAME, Len(oRs!TABLE_NAME) - Len(sDDTableNamePrefix))
            
                Else
            
134                 sTableName = ""
            
                End If
            
136             If Not sTableName = "" Then

138                 If sTableName = sTableNameOld Then    'And Not (oRs!ORDINAL_POSITION = "2" And (Left$(sTableName, 4) = "link")) Then
            
140                     sLocalGUID = GUIDGen
           
142                     .ListItems.Add 1, sLocalGUID, sTableName
144                     .ListItems(sLocalGUID).SubItems(1) = oRs!COLUMN_NAME
     
146                     If InStr(1, sExcludedFields, sTableName & "," & oRs!COLUMN_NAME, vbTextCompare) <> 0 Then

148                         .ListItems(sLocalGUID).Checked = False

                        Else
150                         .ListItems(sLocalGUID).Checked = True

                        End If
                    
                    Else
                    
152                     If (Left$(sTableName, 4) = "link") Then iCountOfLinkedPKs = iCountOfLinkedPKs + 1
                    
                    End If
                    
154                 If Not iCountOfLinkedPKs = 1 Then
156                     sTableNameOld = sTableName
158                     iCountOfLinkedPKs = 0
                    End If
                    
                End If
            
160             oRs.MoveNext
            
            Wend

162         .SortKey = 0    'first column
164         .Sorted = True  'sort it
            On Error Resume Next
166         .ListItems(1).Selected = True   'select the first one
168         .ListItems(1).EnsureVisible     'make sure it is visible

        End With

        '<EhFooter>
        Exit Sub

SetListExcludes_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.SetListExcludes " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function FileExists(sFullPath As String) As Boolean
        '<EhHeader>
        On Error GoTo FileExists_Err
        '</EhHeader>
        Dim oFile As New Scripting.FileSystemObject
100     FileExists = oFile.FileExists(sFullPath)
        '<EhFooter>
        Exit Function

FileExists_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.FileExists " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function GetExcludesAsString() As String
        '<EhHeader>
        On Error GoTo GetExcludesAsString_Err
        '</EhHeader>
            
        Dim sExcludes As String
        Dim itm As ListItem
        Dim i As Integer
        Dim iArray As Integer
100     ReDim ExcludeArray(listExludes.ListItems.Count)
102     i = 1
104     iArray = 1
    
106     Do Until i > listExludes.ListItems.Count
    
108         If listExludes.ListItems(i).Checked = False Then
110             ExcludeArray(iArray).sTableName = listExludes.ListItems(i).Text
112             ExcludeArray(iArray).sFieldName = listExludes.ListItems(i).SubItems(1)
114             iArray = iArray + 1
            End If

116         i = i + 1

        Loop

118     ReDim Preserve ExcludeArray(iArray - 1)
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
120     i = 1
122     GetExcludesAsString = ""

124     Do Until i > UBound(ExcludeArray)
126         GetExcludesAsString = GetExcludesAsString & ExcludeArray(i).sTableName & "," & ExcludeArray(i).sFieldName & ";"
128         i = i + 1
        Loop
            
130     If Len(GetExcludesAsString) > 0 Then GetExcludesAsString = Left(GetExcludesAsString, Len(GetExcludesAsString) - 1)
        
        '<EhFooter>
        Exit Function

GetExcludesAsString_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.GetExcludesAsString " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function GetAccessRightsAsString() As String
        '<EhHeader>
        On Error GoTo GetAccessRightsAsString_Err
        '</EhHeader>

        Dim sExcludes As String

100     GetAccessRightsAsString = ""

102     With RSAccessRights
    
104         If Not .EOF Or Not .Bof Then
        
106             .MoveFirst
    
108             Do Until .EOF
    
110                 GetAccessRightsAsString = GetAccessRightsAsString & .fields(0).Value & ","

112                 If .fields("Read").Value = True Then GetAccessRightsAsString = GetAccessRightsAsString & "r"
114                 If .fields("Append").Value = True Then GetAccessRightsAsString = GetAccessRightsAsString & "a"
116                 If .fields("Edit").Value = True Then GetAccessRightsAsString = GetAccessRightsAsString & "e"
118                 If .fields("Delete").Value = True Then GetAccessRightsAsString = GetAccessRightsAsString & "d"
120                 GetAccessRightsAsString = GetAccessRightsAsString & ";"
122                 .MoveNext
        
                Loop
    
124             .MoveFirst
        
            End If
        
            ' MsgBox GetAccessRightsAsString
        End With
  
        'If Len(GetAccessRightsAsString) > 0 Then GetAccessRightsAsString = Left(GetAccessRightsAsString, Len(GetAccessRightsAsString) - 1)
        
        '<EhFooter>
        Exit Function

GetAccessRightsAsString_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.GetAccessRightsAsString " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub UpdateRSAllData()
        '<EhHeader>
        On Error GoTo UpdateRSAllData_Err
        '</EhHeader>
    
        Dim sExcludes As String
        Dim sAccessRights As String
        Dim bReturnValue As Boolean
    
100     If Not RSDefinedDynDatabases.EOF And Not RSDefinedDynDatabases.Bof Then
    
102         If Not RSAccessRights Is Nothing Then

104             With RSAccessRights
    
106                 If Not .Bof Or Not .EOF Then
                
108                     .MoveFirst
        
110                     Do Until .EOF

112                         If .fields("Append").Value = True Or .fields("Edit").Value = True Or .fields("Delete").Value = True Then
114                             .fields("Read").Value = True
                            End If

116                         .MoveNext
                        Loop
                
118                     .MoveFirst

                    End If

                End With
    
120             sExcludes = ""
122             sAccessRights = ""
124             PromptToSaveDDDatabaseDef = False

126             sExcludes = GetExcludesAsString
128             sAccessRights = GetAccessRightsAsString
        
130             RSDefinedDynDatabases.fields("ExcludedFields").Value = sExcludes
132             RSDefinedDynDatabases.fields("AccessRights").Value = sAccessRights
        
            End If
    
        End If

        '<EhFooter>
        Exit Sub

UpdateRSAllData_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.UpdateRSAllData " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CommitToDatabase()
        '<EhHeader>
        On Error GoTo CommitToDatabase_Err
        '</EhHeader>

100     If listExludes.Visible Then
102         UpdateRSAllData
        End If

104     If MsgBox("Do you want to save this dynamic database definition?", vbYesNo, "Save?") = vbYes Then

            'RSDefinedDynDatabases.Filter = adFilterPendingRecords

106         If Not RSDefinedDynDatabases.EOF Or Not RSDefinedDynDatabases.Bof Then
                
108             RSDefinedDynDatabases.MoveFirst
110             bReturnValue = SaveSilentHttpCommsRS(RSDefinedDynDatabases, sWebsite & "Oasis.asp", True)

112             If bReturnValue Then
114                 IncrementProfileSettingVersion sWebsite, "SettingValue7", sUserGroupName
116                 MsgBox "Data saved to server"
118                 PromptToSaveDDDatabaseDef = True
                Else
120                 MsgBox "Saving to server failed!"
122                 PromptToSaveDDDatabaseDef = False
                End If
        
            End If
        
            'RSDefinedDynDatabases.Filter = adFilterNone
124         Unload Me
        End If
    
        '<EhFooter>
        Exit Sub

CommitToDatabase_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.CommitToDatabase " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSave_Click()
        '<EhHeader>
        On Error GoTo cmdSave_Click_Err
        '</EhHeader>
    
        On Error Resume Next
100     dxDBGridAccessRights.Dataset.Post
102     dxDBGridDefinedDBs.Dataset.Post

        On Error GoTo cmdSave_Click_Err
104     CommitToDatabase
        '<EhFooter>
        Exit Sub

cmdSave_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.cmdSave_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Init(sPassedWebsite As String)
        '<EhHeader>
        On Error GoTo init_Err
        '</EhHeader>

100     sWebsite = WebSite
102     Call SetDefinedDDList
106     LoadListDatas

        '<EhFooter>
        Exit Sub

init_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.Init " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetDefinedDDList()
        '<EhHeader>
        On Error GoTo SetDefinedDDList_Err
        '</EhHeader>

        Dim sString As String
        Dim sGUID As String

100     sString = sWebsite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM [" & sUserGroupName & "DynamicDataDefs] ORDER BY DDDefName")
102     Set RSDefinedDynDatabases = OpenSilentHttpCommsRS(sString, True)
    
104     Set dxDBGridDefinedDBs.DataSource = RSDefinedDynDatabases
106     dxDBGridDefinedDBs.Columns.RetrieveFields
108     dxDBGridDefinedDBs.Columns(0).Visible = False
110     dxDBGridDefinedDBs.Columns(1).Visible = True
112     dxDBGridDefinedDBs.Columns(2).Visible = True
114     dxDBGridDefinedDBs.Columns(3).Visible = False
116     dxDBGridDefinedDBs.Columns(4).Visible = False
118     dxDBGridDefinedDBs.Columns(5).Visible = False
    
120     dxDBGridDefinedDBs.Columns(1).ReadOnly = True
122     dxDBGridDefinedDBs.Columns(2).ReadOnly = False
        
        '<EhFooter>
        Exit Sub

SetDefinedDDList_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.SetDefinedDDList " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

'Private Sub SetAvailableDDList()
'        '<EhHeader>
'        On Error GoTo SetAvailableDDList_Err
'        '</EhHeader>
'
'        Dim oConn As New ADODB.Connection
'        Dim oRs As New ADODB.Recordset
'        Dim sString As String
'        Dim sGUID As String
'        Dim sTableName As String
'        Dim sTableNameOld As String
'
'100     With oConn
'102         .CursorLocation = adUseClient
'104         .ConnectionString = GetConnectionString(CreateAppPath & "\data\db\Oasisclient.mdb")
'106         .open
'        End With
'
'108     Set oRs = oConn.OpenSchema(adSchemaColumns)
'110     oRs.Sort = "ORDINAL_POSITION DESC"
'
'112     Do Until oRs.EOF
'
'114         If Left$(oRs.fields("TABLE_NAME").Value, 2) = "dd" Then
'
'116             sGUID = GUIDGen
'118             sTableName = oRs.fields("TABLE_NAME").Value
'120             sTableName = Right$(sTableName, Len(sTableName) - 3)
'122             sTableName = Left$(sTableName, InStr(1, sTableName, "_", vbTextCompare) - 1)
'
'124             If Not sTableName = sTableNameOld Then
'126                 sAvailableTableNames = sAvailableTableNames & ";" & sTableName
'128                 listAvailable.AddItem sTableName
'130                 sTableNameOld = sTableName
'                End If
'
'            End If
'
'132         oRs.MoveNext
'        Loop
'
'        '<EhFooter>
'        Exit Sub
'
'SetAvailableDDList_Err:
'        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.SetAvailableDDList " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub

Private Sub dxDBGridDefinedDBs_OnBeforeScroll(Allow As Boolean)
        '<EhHeader>
        On Error GoTo dxDBGridDefinedDBs_OnBeforeScroll_Err
        '</EhHeader>
 
100     If listExludes.Visible Then
            On Error Resume Next
102         dxDBGridAccessRights.Dataset.Post
104         UpdateRSAllData
        End If

        '<EhFooter>
        Exit Sub

dxDBGridDefinedDBs_OnBeforeScroll_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.dxDBGridDefinedDBs_OnBeforeScroll " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadListDatas()
        
    Dim sConnectionString As String

    If Not RSDefinedDynDatabases.EOF And Not RSDefinedDynDatabases.Bof Then

        listExludes.Visible = True
        dxDBGridAccessRights.Visible = True
        
        dxLabel3.Visible = True
        dxLabel1(1).Visible = True
        
        sConnectionString = RSDefinedDynDatabases.fields("ConnectionString").Value
        sConnectionString = Replace(sConnectionString, "\data\db\dynamicdata", CreateAppPath & "\data\db\dynamicdata", , , vbTextCompare)

        If Not RemoteConn Is Nothing Then
            If RemoteConn.State = adStateOpen Then RemoteConn.Close
        End If

        Set RemoteConn = New ADODB.Connection
        RemoteConn.ConnectionString = sConnectionString
        RemoteConn.CursorLocation = adUseClient
        RemoteConn.open
                
        SetListExcludes RSDefinedDynDatabases.fields("DDDefName").Value, IIf(IsNull(RSDefinedDynDatabases.fields("ExcludedFields").Value), "", RSDefinedDynDatabases.fields("ExcludedFields").Value)
        SetListAccessRights RSDefinedDynDatabases.fields("DDDefName").Value, IIf(IsNull(RSDefinedDynDatabases.fields("AccessRights").Value), "", RSDefinedDynDatabases.fields("AccessRights").Value)
    
    End If

End Sub

Private Sub dxDBGridDefinedDBs_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, _
                                            ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
        '<EhHeader>
        On Error GoTo dxDBGridDefinedDBs_OnChangeNode_Err
        '</EhHeader>

100     If Not RSDefinedDynDatabases.Status = 4 Then
102         If Not OldNode.Values(0) = Node.Values(0) Then
104             LoadListDatas
            End If
        End If

        '<EhFooter>
        Exit Sub

dxDBGridDefinedDBs_OnChangeNode_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.dxDBGridDefinedDBs_OnChangeNode " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDBGridDefinedDBs_OnDblClick()
        '<EhHeader>
        On Error GoTo dxDBGridDefinedDBs_OnDblClick_Err
        '</EhHeader>

100     If MsgBox("Are you sure you want to delete this record?", vbYesNo, "Confirm deletion") = vbYes Then
    
102         RSDefinedDynDatabases.Filter = RSDefinedDynDatabases.fields(0).Name & " = '" & RSDefinedDynDatabases.fields(0).Value & "'"
104         RSDefinedDynDatabases.Delete adAffectCurrent
106         RSDefinedDynDatabases.Filter = adFilterNone
    
        End If
    
        '<EhFooter>
        Exit Sub

dxDBGridDefinedDBs_OnDblClick_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.dxDBGridDefinedDBs_OnDblClick " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
        'Set Me.Picture = g_PictureDialogSmall
100     Me.WindowState = 2

102     With Me.listExludes       'reference

104         .CheckBoxes = True          'show with check boxes
106         .MultiSelect = True         'allow multi-selection
108         .HideSelection = False      'keep selection when lost focus
110         .FullRowSelect = True       'select the full row
112         .GridLines = True           'show grid lines
114         .View = lvwReport           'show details
116         .ColumnHeaders.Add          'add a couple columns
118         .ColumnHeaders.Add
120         .ColumnHeaders(1).Text = "Table Name"     'name the columns
122         .ColumnHeaders(2).Text = "Field Name"
124         .ColumnHeaders(1).Width = .Width * 0.5  'column width based on width of list view
126         .ColumnHeaders(2).Width = .Width * 0.45
128         .Arrange = lvwAutoLeft

        End With
        
        'cmdAdd.Caption = ">" & Chr(13) & Chr(13) & "A" & Chr(13) & "D" & Chr(13) & "D" & Chr(13) & Chr(13) & ">"

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     Set ofrmMain = Nothing
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.Form_Unload " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub listExludes_BeforeLabelEdit(Cancel As Integer)
        '<EhHeader>
        On Error GoTo listExludes_BeforeLabelEdit_Err
        '</EhHeader>
100     Cancel = True
        '<EhFooter>
        Exit Sub

listExludes_BeforeLabelEdit_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.listExludes_BeforeLabelEdit " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
