VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDynamicDataMenu 
   BackColor       =   &H0050C0A4&
   Caption         =   "Dynamic Data Configuration Wizard"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8760
   Icon            =   "frmDynamicDataMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8760
      _cx             =   15452
      _cy             =   11033
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   2505
         Left            =   90
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   435
         Width           =   8580
         _cx             =   15134
         _cy             =   4419
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
         GridRows        =   1
         GridCols        =   2
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmDynamicDataMenu.frx":68EA
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add new Dynamic Database"
            Height          =   2505
            Left            =   7680
            TabIndex        =   11
            Top             =   0
            Width           =   900
         End
         Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridDefinedDBs 
            Height          =   2505
            Left            =   0
            OleObjectBlob   =   "frmDynamicDataMenu.frx":6930
            TabIndex        =   12
            Top             =   0
            Width           =   7620
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridAccessRights 
         Height          =   2505
         Left            =   5850
         OleObjectBlob   =   "frmDynamicDataMenu.frx":75D8
         TabIndex        =   6
         Top             =   3345
         Width           =   2820
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   5910
         Width           =   3945
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   4470
         TabIndex        =   2
         Top             =   5910
         Width           =   4200
      End
      Begin MSComctlLib.ListView listExludes 
         Height          =   2505
         Left            =   90
         TabIndex        =   1
         Top             =   3345
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   4419
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
      Begin MSComctlLib.ListView listGridExcludes 
         Height          =   2505
         Left            =   2970
         TabIndex        =   7
         Top             =   3345
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   4419
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
      Begin CONTROLSLibCtl.dxLabel dxLabel2 
         Height          =   285
         Left            =   90
         TabIndex        =   9
         Top             =   90
         Width           =   8580
         _Version        =   0
         _cx             =   15134
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
      Begin CONTROLSLibCtl.dxLabel dxLabel4 
         Height          =   285
         Left            =   2970
         TabIndex        =   8
         Top             =   3000
         Width           =   2820
         _Version        =   0
         _cx             =   4974
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
         Caption         =   "Exlcuded from grid"
         BackStyle       =   1
         BackColor       =   5292196
         ForeColor       =   16777215
         LabelStyle      =   0
         Label3dStyle    =   2
         Label3dOrientation=   4
         Label3dDepth    =   0
         PenWidth        =   1
         Angle           =   0
         ShadowColor     =   10526880
      End
      Begin CONTROLSLibCtl.dxLabel dxLabel1 
         Height          =   285
         Index           =   1
         Left            =   5850
         TabIndex        =   5
         Top             =   3000
         Width           =   2820
         _Version        =   0
         _cx             =   4974
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
         Caption         =   "Access rights"
         BackStyle       =   1
         BackColor       =   5292196
         ForeColor       =   16777215
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
         TabIndex        =   4
         Top             =   3000
         Width           =   2820
         _Version        =   0
         _cx             =   4974
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
         Caption         =   "Disabled in data entry"
         BackStyle       =   1
         BackColor       =   5292196
         ForeColor       =   16777215
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
Option Explicit
Private RSAccessRights As ADODB.Recordset
Private RSDefinedDynDatabases As ADODB.Recordset
Private RemoteConn As ADODB.Connection
Private sWebsite As String
Private sUserGroupName As String
Private sUserGroupGUID As String

Private Type TableInfo
    sTableName As String
    sFieldName As String
End Type

Public Sub setUserGroupsRS(sPassedUserGroupName As String, _
                           sPassedUserGroupGUID As String)

    sUserGroupName = sPassedUserGroupName
    sUserGroupGUID = sPassedUserGroupGUID

End Sub



Private Sub C1Elastic1_ResizeChildren()
    With Me.listExludes       'reference
            
        .ColumnHeaders(1).Width = .Width * 0.45 'column width based on width of list view
        .ColumnHeaders(2).Width = .Width * 0.4

    End With
        
    With Me.listGridExcludes       'reference

        .ColumnHeaders(1).Width = .Width * 0.45  'column width based on width of list view
        .ColumnHeaders(2).Width = .Width * 0.4

    End With
End Sub

Private Sub cmdAdd_Click()
        '<EhHeader>
        On Error GoTo cmdAdd_Click_Err
        '</EhHeader>

100     frmDialogWithTwoFields.Caption = "New Dynamic Data Def"
102     frmDialogWithTwoFields.lbl1.Caption = "Please specify the prefix for each Dynamic Data Table" & Chr(13) & "(with no spaces as shown below in the textbox)"
104     frmDialogWithTwoFields.lbl2.Caption = "Please specify the connection string" & Chr(13) & "(please use a relative path as shown below in the textbox and the database named 'DynamicDataDB.MDB' if you are using MSACCESS)"
106     frmDialogWithTwoFields.Height = frmDialogWithTwoFields.Height + (frmDialogWithTwoFields.txt2.Height * 2)
108     frmDialogWithTwoFields.txt2.Height = frmDialogWithTwoFields.txt2.Height * 3
110     frmDialogWithTwoFields.txt1 = "OCHA-NFI-Database"
112     frmDialogWithTwoFields.txt2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\data\db\dynamicdata\DynamicDataDB.MDB;Mode=ReadWrite|Share Deny None;Persist Security Info=False;"
114     frmDialogWithTwoFields.Show vbModal, Me
        
116     If frmDialogWithTwoFields.bClickedOK Then
        
118         If Not frmDialogWithTwoFields.sText1 = "" And Not frmDialogWithTwoFields.sText2 = "" Then
                
120             RSDefinedDynDatabases.AddNew
122             RSDefinedDynDatabases.fields("DynamDBDefGUID").Value = GUIDGen
124             RSDefinedDynDatabases.fields("GroupGUID").Value = sUserGroupGUID
126             RSDefinedDynDatabases.fields("DDDefName").Value = frmDialogWithTwoFields.sText1
128             RSDefinedDynDatabases.fields("Description").Value = "edit this description"
130             RSDefinedDynDatabases.fields("ConnectionString").Value = frmDialogWithTwoFields.sText2
            
            End If
        
        End If
        
        '<EhFooter>
        Exit Sub

cmdAdd_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.cmdAdd_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
    Unload Me
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
        Dim sTableName As String
        
        Dim oDB As ADOX.Catalog
        Dim itbl As ADOX.Table
 
100     Set oDB = New ADOX.Catalog
102     Set itbl = New ADOX.Table
104     Set oDB.ActiveConnection = RemoteConn
     
106     Set RSAccessRights = New ADODB.Recordset
108     RSAccessRights.fields.Append "Table Name", adVarChar, 50
110     RSAccessRights.fields.Append "Read", adBoolean
112     RSAccessRights.fields.Append "Append", adBoolean
114     RSAccessRights.fields.Append "Edit", adBoolean
116     RSAccessRights.fields.Append "Delete", adBoolean
118     RSAccessRights.open
    
120     sLocalGUID = GUIDGen
        
122     sDDTableNamePrefix = "dd_" & sDDDefName & "_"

124     For Each itbl In oDB.Tables
         
126         If itbl.Properties("Jet OLEDB:Create Link") = True Then

128             sTableName = ""
130         ElseIf (Left$(itbl.Name, Len(sDDTableNamePrefix) + 4) = sDDTableNamePrefix & "link") Then
            
132             sTableName = Right(itbl.Name, Len(itbl.Name) - Len(sDDTableNamePrefix))
            
134         ElseIf itbl.Name = sDDTableNamePrefix & "mastertable" Then
            
136             sTableName = "mastertable"
            
138         ElseIf (Left$(itbl.Name, Len(sDDTableNamePrefix) + 2) = sDDTableNamePrefix & "dd") Then
            
140             sTableName = Right(itbl.Name, Len(itbl.Name) - Len(sDDTableNamePrefix))
            
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

        Next
         
188     Set dxDBGridAccessRights.DataSource = RSAccessRights
190     dxDBGridAccessRights.Columns.RetrieveFields
192     dxDBGridAccessRights.Columns(0).ReadOnly = True
194     dxDBGridAccessRights.Columns(1).Width = 60
196     dxDBGridAccessRights.Columns(2).Width = 60
198     dxDBGridAccessRights.Columns(3).Width = 60
200     dxDBGridAccessRights.Columns(4).Width = 60

202     If Not RSAccessRights.RecordCount > 0 Then
204         MsgBox "Please check your DDDefName parameter.  It appears to be in error", vbInformation
206         listExludes.Visible = False
208         listGridExcludes.Visible = False
210         dxDBGridAccessRights.Visible = False
212         dxLabel3.Visible = False
214         dxLabel1(1).Visible = False
216         dxLabel4.Visible = False
        End If

218     Set oDB = Nothing
220     Set itbl = Nothing
        '<EhFooter>
        Exit Sub

SetListAccessRights_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.SetListAccessRights " & "at line " & Erl
        'Resume Next
        '</EhFooter>
End Sub

Private Function GetTableNames(oConn As ADODB.Connection, _
                               sTablePrefix As String)
        '<EhHeader>
        On Error GoTo GetTableNames_Err
        '</EhHeader>
 
        Dim oDB As ADOX.Catalog
        Dim itbl As ADOX.Table
 
100     Set oDB = New ADOX.Catalog
102     Set itbl = New ADOX.Table
104     Set oDB.ActiveConnection = oConn
106     GetTableNames = ""

108     For Each itbl In oDB.Tables

110         If InStr(1, itbl.Name, sTablePrefix) > 0 Then

112             If Not itbl.Properties("Jet OLEDB:Create Link") = True Then
                    'get linked table
114                 GetTableNames = GetTableNames & itbl.Name & ","
                End If
            End If

        Next
        
116     Set oDB = Nothing
118     Set itbl = Nothing
120     If Len(GetTableNames) > 0 Then GetTableNames = Left$(GetTableNames, Len(GetTableNames) - 1)

        '<EhFooter>
        Exit Function

GetTableNames_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.GetTableNames " & _
               "at line " & Erl
        'Resume Next
        '</EhFooter>
End Function

Private Sub SetListExcludes(sDDDefName As String, _
                            sExcludedFields As String, _
                            sExcludedGrid As String)
        '<EhHeader>
        On Error GoTo SetListExcludes_Err
        '</EhHeader>
    
        Dim iINx() As Integer
        Dim sLocalGUID As String
        Dim sDDTableNamePrefix As String
        Dim sTableNameList As String
        
        Dim oRs As New ADODB.Recordset
        Dim sTableNameOld As String
        Dim sTableName As String
        Dim iCountOfLinkedPKs As Integer
        
        sTableNameList = GetTableNames(RemoteConn, "dd_" & sDDDefName & "_")
        
    
100     Set oRs = RemoteConn.OpenSchema(adSchemaColumns)
102     oRs.Sort = "TABLE_NAME, ORDINAL_POSITION"

104     listExludes.ListItems.Clear
106     listGridExcludes.ListItems.Clear

108     With Me.listExludes

110         sLocalGUID = GUIDGen
        
112         sDDTableNamePrefix = "dd_" & sDDDefName & "_"
        
114         While Not oRs.EOF
         
116             If (Left$(oRs!TABLE_NAME, Len(sDDTableNamePrefix) + 4) = sDDTableNamePrefix & "link") Then
            
118                 sTableName = Right(oRs!TABLE_NAME, Len(oRs!TABLE_NAME) - Len(sDDTableNamePrefix))
            
120             ElseIf oRs!TABLE_NAME = sDDTableNamePrefix & "mastertable" Then
            
122                 sTableName = "mastertable"
            
124             ElseIf (Left$(oRs!TABLE_NAME, Len(sDDTableNamePrefix) + 2) = sDDTableNamePrefix & "dd") Then
            
126                 sTableName = Right(oRs!TABLE_NAME, Len(oRs!TABLE_NAME) - Len(sDDTableNamePrefix))
            
                Else
            
128                 sTableName = ""
            
                End If
                
                If Not InStr(1, sTableNameList, sTableName, vbTextCompare) > 0 Then
                    sTableName = ""
                End If
                
                If Right$(sTableName, 4) = "_GEO" Then sTableName = ""
            
130             If Not sTableName = "" Then

132                 'If sTableName = sTableNameOld Then    'And Not (oRs!ORDINAL_POSITION = "2" And (Left$(sTableName, 4) = "link")) Then
            
134                     sLocalGUID = GUIDGen
                     
136                     listGridExcludes.ListItems.Add 1, sLocalGUID, sTableName
138                     listGridExcludes.ListItems(sLocalGUID).SubItems(1) = oRs.fields("COLUMN_NAME").Value
                        
140                     .ListItems.Add 1, sLocalGUID, sTableName
142                     .ListItems(sLocalGUID).SubItems(1) = oRs.fields("COLUMN_NAME").Value
     
144                     If InStr(1, sExcludedGrid, sTableName & "," & oRs.fields("COLUMN_NAME").Value, vbTextCompare) <> 0 Then
146                         listGridExcludes.ListItems(sLocalGUID).Checked = False
                        Else
148                         listGridExcludes.ListItems(sLocalGUID).Checked = True
                        End If
                        
150                     If InStr(1, sExcludedFields, sTableName & "," & oRs.fields("COLUMN_NAME").Value, vbTextCompare) <> 0 Then
152                         .ListItems(sLocalGUID).Checked = False
                        Else
154                         .ListItems(sLocalGUID).Checked = True
                        End If
                    
                   ' Else
                    
156                   '  If (Left$(sTableName, 4) = "link") Then iCountOfLinkedPKs = iCountOfLinkedPKs + 1
                    
                '    End If
                    
158                 If Not iCountOfLinkedPKs = 1 Then
160                     sTableNameOld = sTableName
162                     iCountOfLinkedPKs = 0
                    End If
                    
                End If
            
164             oRs.MoveNext
            
            Wend

166         listGridExcludes.SortKey = 0
168         listGridExcludes.Sorted = True
170         .SortKey = 0    'first column
172         .Sorted = True  'sort it
            On Error Resume Next
174         .ListItems(1).Selected = True   'select the first one
176         .ListItems(1).EnsureVisible     'make sure it is visible
178         listGridExcludes.ListItems(1).Selected = True
180         listGridExcludes.ListItems(1).EnsureVisible

        End With

        '<EhFooter>
        Exit Sub

SetListExcludes_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.SetListExcludes " & _
               "at line " & Erl
       ' Resume Next
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
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.FileExists " & _
               "at line " & Erl
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
        Dim ExcludeArray() As TableInfo
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
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.GetExcludesAsString " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function GetGridExcludesAsString() As String
        '<EhHeader>
        On Error GoTo GetGridExcludesAsString_Err
        '</EhHeader>
            
        Dim sExcludes As String
        Dim itm As ListItem
        Dim i As Integer
        Dim iArray As Integer
        Dim ExcludeGridArray() As TableInfo
100     ReDim ExcludeGridArray(listExludes.ListItems.Count)
    
102     i = 1
104     iArray = 1

106     Do Until i > listGridExcludes.ListItems.Count
    
108         If listGridExcludes.ListItems(i).Checked = False Then
110             ExcludeGridArray(iArray).sTableName = listGridExcludes.ListItems(i).Text
112             ExcludeGridArray(iArray).sFieldName = listGridExcludes.ListItems(i).SubItems(1)
114             iArray = iArray + 1
            End If

116         i = i + 1
        Loop

118     ReDim Preserve ExcludeGridArray(iArray - 1)
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
120     i = 1
122     GetGridExcludesAsString = ""

124     Do Until i > UBound(ExcludeGridArray)
126         GetGridExcludesAsString = GetGridExcludesAsString & ExcludeGridArray(i).sTableName & "," & ExcludeGridArray(i).sFieldName & ";"
128         i = i + 1
        Loop
            
130     If Len(GetGridExcludesAsString) > 0 Then GetGridExcludesAsString = Left(GetGridExcludesAsString, Len(GetGridExcludesAsString) - 1)
        
        '<EhFooter>
        Exit Function

GetGridExcludesAsString_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.GetGridExcludesAsString " & _
               "at line " & Erl
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
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.GetAccessRightsAsString " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub UpdateRSAllData()
        '<EhHeader>
        On Error GoTo UpdateRSAllData_Err
        '</EhHeader>
    
        Dim sExcludes As String
        Dim sGridExcludes As String
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
122             sGridExcludes = ""
124             sAccessRights = ""
                'PromptToSaveDDDatabaseDef = False

126             sExcludes = GetExcludesAsString
128             sGridExcludes = GetGridExcludesAsString
130             sAccessRights = GetAccessRightsAsString
        
132             RSDefinedDynDatabases.fields("ExcludedFields").Value = sExcludes
134             RSDefinedDynDatabases.fields("ExcludedGrid").Value = sGridExcludes
136             RSDefinedDynDatabases.fields("AccessRights").Value = sAccessRights
        
            End If
    
        End If

        '<EhFooter>
        Exit Sub

UpdateRSAllData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.UpdateRSAllData " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CommitToDatabase()
        '<EhHeader>
        On Error GoTo CommitToDatabase_Err
        '</EhHeader>
        
        Dim bReturnValue As String

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
                    'PromptToSaveDDDatabaseDef = True
                Else
118                 MsgBox "Saving to server failed!"
                    'PromptToSaveDDDatabaseDef = False
                End If
        
            End If
        
            'RSDefinedDynDatabases.Filter = adFilterNone
120         Unload Me
        End If
    
        '<EhFooter>
        Exit Sub

CommitToDatabase_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.CommitToDatabase " & _
               "at line " & Erl
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
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.cmdSave_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Init(sPassedWebsite As String)
        '<EhHeader>
        On Error GoTo init_Err
        '</EhHeader>

100     sWebsite = WebSite
102     Call SetDefinedDDList
104     LoadListDatas

        '<EhFooter>
        Exit Sub

init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.Init " & _
               "at line " & Erl
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
120     dxDBGridDefinedDBs.Columns(6).Visible = False
122     dxDBGridDefinedDBs.Columns(7).Visible = True
124     dxDBGridDefinedDBs.Columns(8).Visible = True
126     dxDBGridDefinedDBs.Columns(1).Width = 200
128     dxDBGridDefinedDBs.Columns(2).Width = 200
130     dxDBGridDefinedDBs.Columns(7).Width = 500
132     dxDBGridDefinedDBs.Columns(8).Width = 50
        
        '<EhFooter>
        Exit Sub

SetDefinedDDList_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.SetDefinedDDList " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

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
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.dxDBGridDefinedDBs_OnBeforeScroll " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadListDatas()
        '<EhHeader>
        On Error GoTo LoadListDatas_Err
        '</EhHeader>
        
        Dim sConnectionString As String

100     If Not RSDefinedDynDatabases.EOF And Not RSDefinedDynDatabases.Bof Then

102         listExludes.Visible = True
104         listGridExcludes.Visible = True
106         dxDBGridAccessRights.Visible = True
        
108         dxLabel3.Visible = True
110         dxLabel1(1).Visible = True
112         dxLabel4.Visible = True
114         sConnectionString = RSDefinedDynDatabases.fields("ConnectionString").Value
116         sConnectionString = Replace(sConnectionString, "\data\db\dynamicdata", CreateAppPath & "\data\db\dynamicdata", , , vbTextCompare)

118         If Not RemoteConn Is Nothing Then
120             If RemoteConn.State = adStateOpen Then RemoteConn.Close
            End If

122         Set RemoteConn = New ADODB.Connection
124         RemoteConn.ConnectionString = sConnectionString
126         RemoteConn.CursorLocation = adUseClient
            
            On Error GoTo ErrorInConnStr
128         RemoteConn.open
            On Error GoTo LoadListDatas_Err
                
130         SetListExcludes RSDefinedDynDatabases.fields("DDDefName").Value, IIf(IsNull(RSDefinedDynDatabases.fields("ExcludedFields").Value), "", RSDefinedDynDatabases.fields("ExcludedFields").Value), IIf(IsNull(RSDefinedDynDatabases.fields("ExcludedGrid").Value), "", RSDefinedDynDatabases.fields("ExcludedGrid").Value)
132         SetListAccessRights RSDefinedDynDatabases.fields("DDDefName").Value, IIf(IsNull(RSDefinedDynDatabases.fields("AccessRights").Value), "", RSDefinedDynDatabases.fields("AccessRights").Value)
    
        End If

        Exit Sub
        
ErrorInConnStr:
        
134     MsgBox "There is an error in the connection string!", vbInformation
136     listExludes.Visible = False
138     listGridExcludes.Visible = False
140     dxDBGridAccessRights.Visible = False
142     dxLabel3.Visible = False
144     dxLabel1(1).Visible = False
146     dxLabel4.Visible = False
        Exit Sub
        
        '<EhFooter>
        Exit Sub

LoadListDatas_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDynamicDataMenu.LoadListDatas " & "at line " & Erl
        '</EhFooter>
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
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.dxDBGridDefinedDBs_OnChangeNode " & _
               "at line " & Erl
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
108         LoadListDatas

        End If
    
        '<EhFooter>
        Exit Sub

dxDBGridDefinedDBs_OnDblClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.dxDBGridDefinedDBs_OnDblClick " & _
               "at line " & Erl
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
124         .ColumnHeaders(1).Width = .Width * 0.5 * 1.75 'column width based on width of list view
126         .ColumnHeaders(2).Width = .Width * 0.45 * 1.75
128         .Arrange = lvwAutoLeft

        End With
        
130     With Me.listGridExcludes       'reference

132         .CheckBoxes = True          'show with check boxes
134         .MultiSelect = True         'allow multi-selection
136         .HideSelection = False      'keep selection when lost focus
138         .FullRowSelect = True       'select the full row
140         .GridLines = True           'show grid lines
142         .View = lvwReport           'show details
144         .ColumnHeaders.Add          'add a couple columns
146         .ColumnHeaders.Add
148         .ColumnHeaders(1).Text = "Table Name"     'name the columns
150         .ColumnHeaders(2).Text = "Field Name"
152         .ColumnHeaders(1).Width = .Width * 0.5 * 1.75 'column width based on width of list view
154         .ColumnHeaders(2).Width = .Width * 0.45 * 1.75
156         .Arrange = lvwAutoLeft

        End With
        
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmDynamicDataMenu.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Paint()
'MsgBox "hh"
End Sub

Private Sub Form_Resize()

'MsgBox "dude"


        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    Set RSAccessRights = Nothing
    Set RSDefinedDynDatabases = Nothing
    Set RemoteConn = Nothing
    
End Sub

Private Sub listExludes_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub listGridExcludes_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub
