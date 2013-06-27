VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmSynchLayerWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Synchronisation Wizard"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7320
   Icon            =   "frmSynchLayerWizard.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   6360
      TabIndex        =   35
      Top             =   4920
      Width           =   915
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "Admin Tools"
      Height          =   285
      Left            =   6210
      TabIndex        =   34
      Top             =   30
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   285
      Left            =   3480
      TabIndex        =   16
      Top             =   4920
      Width           =   915
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   285
      Left            =   4440
      TabIndex        =   15
      Top             =   4920
      Width           =   915
   End
   Begin C1SizerLibCtl.C1Tab C1TTabSynchLyr 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   1650
      Width           =   7335
      _cx             =   12938
      _cy             =   5636
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
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "Tab&1|Tab&2|Tab&3|New Tab|New Tab|New Tab"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   2880
         Left            =   15
         OleObjectBlob   =   "frmSynchLayerWizard.frx":6852
         TabIndex        =   29
         Top             =   300
         Width           =   7305
      End
      Begin VB.Frame fraSteaps 
         BorderStyle     =   0  'None
         Height          =   2880
         Index           =   6
         Left            =   9150
         TabIndex        =   7
         Top             =   300
         Width           =   7305
         Begin VB.TextBox txtSessionGUID 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   32
            Top             =   2280
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.TextBox txtOwnerGUID 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1440
            TabIndex        =   31
            Top             =   1680
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.TextBox txtUserGroup 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2280
            TabIndex        =   28
            Top             =   600
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.TextBox txtSynchFreq 
            Height          =   315
            Left            =   2280
            TabIndex        =   24
            Top             =   180
            Width           =   2865
         End
         Begin VB.Label lblSteps 
            AutoSize        =   -1  'True
            Caption         =   "Owner:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   26
            Top             =   660
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label lblSteps 
            AutoSize        =   -1  'True
            Caption         =   "Synch frequency (seconds):"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   25
            Top             =   270
            Width           =   1980
         End
      End
      Begin VB.Frame fraSteaps 
         Caption         =   "Frame1"
         Height          =   2880
         Index           =   4
         Left            =   8850
         TabIndex        =   5
         Top             =   300
         Width           =   7305
         Begin VB.Frame fraSteaps 
            BorderStyle     =   0  'None
            Height          =   4800
            Index           =   5
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   7245
            Begin VB.CheckBox chkWriteAccess 
               Caption         =   "Write Access"
               DataField       =   "AllowWrite"
               BeginProperty DataFormat 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "True"
                  FalseValue      =   "False"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
               DataSource      =   "RSSynchTable"
               Height          =   315
               Left            =   5400
               TabIndex        =   30
               Top             =   240
               Width           =   1245
            End
            Begin VB.CheckBox chkAutoUpdate 
               Caption         =   "Forced Update"
               DataField       =   "AutoUpdate"
               BeginProperty DataFormat 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "True"
                  FalseValue      =   "False"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
               DataSource      =   "RSSynchTable"
               Height          =   315
               Left            =   3570
               TabIndex        =   27
               Top             =   210
               Width           =   1725
            End
            Begin VB.CheckBox chkIsActive 
               Caption         =   "Is Active"
               DataField       =   "IsActive"
               BeginProperty DataFormat 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "True"
                  FalseValue      =   "False"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
               DataSource      =   "RSSynchTable"
               Height          =   285
               Left            =   2520
               TabIndex        =   22
               Top             =   210
               Width           =   1125
            End
            Begin VB.CheckBox chkInLegend 
               Caption         =   "In Legend"
               DataField       =   "InLegend"
               BeginProperty DataFormat 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "True"
                  FalseValue      =   "False"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
               DataSource      =   "RSSynchTable"
               Height          =   285
               Left            =   1320
               TabIndex        =   21
               Top             =   210
               Width           =   1275
            End
            Begin VB.CheckBox chkIsGeolayer 
               Caption         =   "Is Geolayer"
               DataField       =   "isGeoTable"
               BeginProperty DataFormat 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "True"
                  FalseValue      =   "False"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
               DataSource      =   "RSSynchTable"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   1155
            End
         End
      End
      Begin VB.Frame fraSteaps 
         BorderStyle     =   0  'None
         Height          =   2880
         Index           =   3
         Left            =   8550
         TabIndex        =   4
         Top             =   300
         Width           =   7305
         Begin VB.TextBox txtDescription 
            Height          =   915
            Left            =   1020
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   180
            Width           =   6105
         End
         Begin VB.Label lblSteps 
            AutoSize        =   -1  'True
            Caption         =   "Description:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   10
            Top             =   270
            Width           =   840
         End
      End
      Begin VB.Frame fraSteaps 
         Caption         =   "Frame1"
         Height          =   2880
         Index           =   1
         Left            =   8250
         TabIndex        =   2
         Top             =   300
         Width           =   7305
         Begin VB.Frame fraSteaps 
            BorderStyle     =   0  'None
            Height          =   4800
            Index           =   2
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   7245
            Begin VB.TextBox txtTable 
               Height          =   375
               Left            =   3600
               TabIndex        =   33
               TabStop         =   0   'False
               Text            =   "Text1"
               Top             =   2160
               Visible         =   0   'False
               Width           =   2175
            End
            Begin VB.TextBox txtLgdCaption 
               Height          =   285
               Left            =   5100
               TabIndex        =   23
               Top             =   510
               Width           =   1995
            End
            Begin VB.TextBox txtLyrAlias 
               Height          =   285
               Left            =   1110
               TabIndex        =   17
               Top             =   540
               Width           =   2685
            End
            Begin VB.ComboBox ComSynchTable 
               Height          =   315
               Left            =   1110
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   120
               Width           =   5955
            End
            Begin VB.Label lblSteps 
               AutoSize        =   -1  'True
               Caption         =   "Table Alias:"
               Height          =   195
               Index           =   4
               Left            =   210
               TabIndex        =   18
               Top             =   540
               Width           =   825
            End
            Begin VB.Label lblSteps 
               AutoSize        =   -1  'True
               Caption         =   "Select Table:"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   12
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblSteps 
               AutoSize        =   -1  'True
               Caption         =   "Legend Caption:"
               Height          =   195
               Index           =   2
               Left            =   3870
               TabIndex        =   9
               Top             =   600
               Width           =   1170
            End
         End
      End
      Begin VB.Frame fraSteaps 
         BorderStyle     =   0  'None
         Height          =   2880
         Index           =   0
         Left            =   7950
         TabIndex        =   1
         Top             =   300
         Width           =   7305
         Begin VB.CommandButton cmdSelectDB 
            Caption         =   "..."
            Height          =   285
            Left            =   6720
            TabIndex        =   14
            Top             =   270
            Width           =   435
         End
         Begin VB.TextBox txtMDB 
            Height          =   285
            Left            =   1320
            TabIndex        =   13
            Top             =   270
            Width           =   5385
         End
         Begin VB.Label lblSteps 
            AutoSize        =   -1  'True
            Caption         =   "Select Database:"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   360
            Width           =   1230
         End
      End
   End
End
Attribute VB_Name = "frmSynchLayerWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSLocalUserGroups As New ADODB.Recordset
Dim RSSynchTable As New ADODB.Recordset
Dim RSTableToSynch As New ADODB.Recordset
Dim bEditEntry As VbMsgBoxResult

Public Sub setUserGroupsRS(ByVal PassedRS As ADODB.Recordset)
    
    Set RSLocalUserGroups = PassedRS
    
End Sub

Private Sub cmdBack_Click()
        '<EhHeader>
        On Error GoTo cmdBack_Click_Err
        '</EhHeader>
    
100     With C1TTabSynchLyr
            
102         If Not .CurrTab = 0 Then
104             .CurrTab = .CurrTab - 1
            End If

        End With
    
        '<EhFooter>
        Exit Sub

cmdBack_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSynchLayerWizard.cmdBack_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
        '<EhHeader>
        On Error GoTo cmdNext_Click_Err
        '</EhHeader>
    
        Dim bVer As Boolean
100     cmdBack.Caption = "Back"

102     With C1TTabSynchLyr
    
104         Select Case .CurrTab
    
                Case 0
            
106                 bVer = True
108                 Me.txtMDB = getClientDBPath & "\OasisClient.mdb"

                    bEditEntry = vbNo

                    If Not RSSynchTable.EOF Or Not RSSynchTable.Bof Then
                    
                        bEditEntry = MsgBox("Do you want to edit the selected entry?  Clicking 'No' will add a new entry", vbYesNoCancel, "Confirm operation")

                        If bEditEntry = vbYes Then
                            Call EditExistingEntry
                            .CurrTab = .CurrTab + 1
                        ElseIf bEditEntry = vbNo Then
                            Call PrepareForNewEntry
                        Else
                            bVer = False
                        End If

                    Else
                        Call PrepareForNewEntry
                    End If

110             Case 1

112                 If Len(txtMDB.Text) > 6 Then

                        If GetTables Then bVer = True
116                     Call PrepareForNewEntry
118                     txtUserGroup.Text = RSLocalUserGroups!Name
120                     txtOwnerGUID.Text = RSLocalUserGroups!sGUID
122                     txtSessionGUID.Text = GUIDGen()

                    Else
124                     MsgBox "You have to choose a DB before continuing..."
                    End If

126             Case 2
               
128                 If Len(txtLyrAlias.Text) < 2 Then
130                     MsgBox "The Alias Name is too short..."
                        Exit Sub
                    End If
                
132                 If Len(txtLgdCaption.Text) < 2 Then
134                     MsgBox "The Caption is too short..."
                        Exit Sub
                    End If
               
136                 bVer = True
138                 txtTable.Text = ComSynchTable.Text
               
140             Case 3

142                 If Len(txtDescription.Text) < 2 Then
144                     MsgBox "The Description is too short..."
                        Exit Sub
                    End If

146                 bVer = True

148             Case 4

150                 bVer = True

152             Case 5

154                 If Not IsNumeric(txtSynchFreq.Text) Then
156                     MsgBox "The Sequence is not numeric..."
                        Exit Sub
                    End If
                
158                 If MsgBox("Do you want to save this info?", vbYesNo, "Confirm Addition") = vbYes Then
                        
160                     If bEditEntry = vbNo Then RSSynchTable.AddNew

162                     With RSSynchTable.fields

164                         .Item("sTableName").Value = ComSynchTable
166                         .Item("sName").Value = txtLyrAlias
168                         .Item("sCaption").Value = txtLgdCaption
170                         .Item("sDescription").Value = txtDescription
172                         .Item("SynchFrequency").Value = txtSynchFreq
174                         .Item("OwnerID").Value = txtOwnerGUID
176                         .Item("sGUID").Value = txtSessionGUID

178                         .Item(chkIsActive.DataField).Value = IIf(chkIsActive.Value = vbChecked, True, False)
180                         .Item(chkInLegend.DataField).Value = IIf(chkInLegend.Value = vbChecked, True, False)
182                         .Item(chkIsGeolayer.DataField).Value = IIf(chkIsGeolayer.Value = vbChecked, True, False)
184                         .Item(chkWriteAccess.DataField).Value = IIf(chkWriteAccess.Value = vbChecked, True, False)
186                         .Item(chkAutoUpdate.DataField).Value = IIf(chkAutoUpdate.Value = vbChecked, True, False)
                        
                        End With
                        
188                     SaveDataToServer ComSynchTable
190                     Unload Me

                        Exit Sub
                    End If
                
192                 bVer = False
            
            End Select

194         If bVer Then
196             If Not .CurrTab = .NumTabs Then
198                 .CurrTab = .CurrTab + 1
                End If
            End If
     
        End With
    
        '<EhFooter>
        Exit Sub

cmdNext_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmSynchLayerWizard.cmdNext_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub EditExistingEntry()
        '<EhHeader>
        On Error GoTo EditExistingEntry_Err
        '</EhHeader>
    
        Dim i As Integer
    
100     With RSSynchTable.fields

102         GetTables

            On Error GoTo TableMissing
        
104         ComSynchTable.Text = .Item("sTableName").Value

            On Error GoTo EditExistingEntry_Err
            
106         txtLyrAlias = .Item("sName").Value
108         txtLgdCaption = .Item("sCaption").Value
110         txtDescription = .Item("sDescription").Value
112         txtSynchFreq = .Item("SynchFrequency").Value
114         txtOwnerGUID = .Item("OwnerID").Value
116         txtSessionGUID = .Item("sGUID").Value

118         chkIsActive.Value = IIf(.Item(chkIsActive.DataField).Value = True, vbChecked, vbUnchecked)
120         chkInLegend.Value = IIf(.Item(chkInLegend.DataField).Value = True, vbChecked, vbUnchecked)
122         chkIsGeolayer.Value = IIf(.Item(chkIsGeolayer.DataField).Value = True, vbChecked, vbUnchecked)
124         chkWriteAccess.Value = IIf(.Item(chkWriteAccess.DataField).Value = True, vbChecked, vbUnchecked)
126         chkAutoUpdate.Value = IIf(.Item(chkAutoUpdate.DataField).Value = True, vbChecked, vbUnchecked)
                     
        End With
 
        '<EhFooter>
        Exit Sub
        
TableMissing:
        
        MsgBox "The table [" & RSSynchTable.fields.Item("sName").Value & "] does not exist in the OASIS Client database.", vbCritical
        Resume Next

EditExistingEntry_Err:

        Resume Next
        '</EhFooter>
End Sub

Private Sub PrepareForNewEntry()
        '<EhHeader>
        On Error GoTo PrepareForNewEntry_Err
        '</EhHeader>

100     Me.chkAutoUpdate.Value = vbUnchecked
102     Me.chkInLegend.Value = vbUnchecked
104     Me.chkIsActive.Value = vbUnchecked
106     Me.chkIsGeolayer.Value = vbUnchecked
108     Me.chkWriteAccess.Value = vbUnchecked
110     Me.txtOwnerGUID = RSLocalUserGroups!sGUID
112     Me.txtSessionGUID = GUIDGen()

        '<EhFooter>
        Exit Sub

PrepareForNewEntry_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSynchLayerWizard.PrepareForNewEntry " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetTables() As Boolean
        '<EhHeader>
        On Error GoTo GetTables_Err
        '</EhHeader>
        Dim cn As New ADODB.Connection
        
100     cn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtMDB.Text & ";"
102     ComSynchTable.Clear

104     ShowAllTables cn, False, ComSynchTable
106     ComSynchTable.ListIndex = 0

108     GetTables = True
        ' Call setDataSources
110     cn.Close
112     Set cn = Nothing
        
        '<EhFooter>
        Exit Function

GetTables_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSynchLayerWizard.GetTables " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub SaveDataToServer(sPassedTableName As String)
        '<EhHeader>
        On Error GoTo SaveDataToServer_Err
        '</EhHeader>

        Dim cn As New ADODB.Connection
        Dim strQuery As String
        Dim bReturnValue As Boolean
        Dim sReturnValue As String
    
100     RSSynchTable.Filter = adFilterPendingRecords

102     If Not RSSynchTable.EOF And Not RSSynchTable.Bof Then
        
104         bReturnValue = m_frmOASISProgress.SaveHttpCommsRS(RSSynchTable, WebSite & "Oasis.asp", True)
            
106         If bReturnValue Then
        
108             IncrementProfileSettingVersion WebSite, "SettingValue8", RSLocalUserGroups.fields("Name").Value
110             MsgBox "Data saved to server"

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Load local version of table
    
112             cn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtMDB.Text & ";"
114             strQuery = "SELECT * FROM " & RSLocalUserGroups!Name & sPassedTableName ' Dunno if this is needed
116             Set RSTableToSynch = New ADODB.Recordset

118             With RSTableToSynch
120                 Set .ActiveConnection = cn
122                 .CursorType = adOpenDynamic
124                 .LockType = adLockBatchOptimistic
126                 .Source = strQuery ' Dunno if this is needed
128                 .CursorLocation = adUseClient
130                 .open "SELECT * FROM " & sPassedTableName, cn
                End With
            
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Load server version of table
    
132             strQuery = WebSite & "Oasis.asp?id=" & CheckEncrypt("SELECT * FROM " & RSLocalUserGroups!Name & sPassedTableName)
134             sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(strQuery, True)

136             If sReturnValue = "-1" Then ' This does not exist then
138                 strQuery = WebSite & "Oasis.asp?checkTbl=" & CheckEncrypt(RSLocalUserGroups!Name & sPassedTableName)
140                 m_frmOASISProgress.SaveHttpCommsRS RSTableToSynch, strQuery, True
                Else
142                 strQuery = WebSite & "Oasis.asp"
144                 m_frmOASISProgress.SaveHttpCommsRS RSTableToSynch, strQuery, True
                End If

            Else
146             MsgBox "Save failed!", vbCritical
            End If
        
        End If

        '<EhFooter>
        Exit Sub

SaveDataToServer_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSynchLayerWizard.SaveDataToServer " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function getClientDBPath()
        '<EhHeader>
        On Error GoTo getClientDBPath_Err
        '</EhHeader>
        
100     getClientDBPath = CreateAppPath & "\Data\DB"

        '<EhFooter>
        Exit Function

getClientDBPath_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSynchLayerWizard.getClientDBPath " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdSelectDB_Click()
        '<EhHeader>
        On Error GoTo cmdSelectDB_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog
    
        On Error Resume Next
100     c.DefaultExt = "mdb"
102     c.Filter = "*.mdb|*.MDB"
104     c.DialogTitle = "OASIS Synch Wizard..."
106     c.InitDir = getClientDBPath
        ' c.Filename = "OasisClient"
108     c.ShowOpen
        
110     txtMDB.Text = c.fileName

        '<EhFooter>
        Exit Sub

cmdSelectDB_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSynchLayerWizard.cmdSelectDB_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTools_Click()
        '<EhHeader>
        On Error GoTo cmdTools_Click_Err
        '</EhHeader>
        Dim m_frmAdminTools As frmAdminTools
100     Set m_frmAdminTools = New frmAdminTools
102     m_frmAdminTools.Show vbModeless, Me
104     Set m_frmAdminTools = Nothing
        '<EhFooter>
        Exit Sub

cmdTools_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSynchLayerWizard.cmdTools_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComSynchTable_LostFocus()
        '<EhHeader>
        On Error GoTo ComSynchTable_LostFocus_Err
        '</EhHeader>
100     txtTable.Text = ComSynchTable.Text
        '<EhFooter>
        Exit Sub

ComSynchTable_LostFocus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSynchLayerWizard.ComSynchTable_LostFocus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDBGrid1_OnDblClick()
    
    Dim sReturn As String
    Dim strQuery As String
    Dim sTableName As String
    
    sTableName = RSSynchTable.fields("sTableName")
    
    If DeleteRecordFromRSAndSave(RSSynchTable, "SettingValue8", RSLocalUserGroups.fields("Name").Value) Then
    
        strQuery = WebSite & "Oasis.asp?id=" & CheckEncrypt("DROP TABLE " & RSLocalUserGroups!Name & sTableName)
        sReturn = m_frmOASISProgress.OpenHttpCommsResponse(strQuery, True)
    
        'TO DO: this deletes the synchtable from the server but we also need to setup protocol for deleting the synch'
        '       table from all the client databases - currently this can only be done via autoupdate
    
        If sReturn = "" Then
            MsgBox "Table [" & RSLocalUserGroups!Name & sTableName & "] successfully deleted on the server"
        Else
            MsgBox "Table [" & RSLocalUserGroups!Name & sTableName & "] is not on the server and hence did not require deletion."
        End If
   
    End If

End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
        Dim sString As String
100     Me.Picture = g_PictureDialogSmall

        C1TTabSynchLyr.TabHeight = 1
        C1TTabSynchLyr.CurrTab = 0

102     DoEvents

104     m_frmDebug.DebugPrint "Getting Synchronisation Table for user: " & txtUserGroup.Text
106     Set RSSynchTable = New ADODB.Recordset
108     sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & RSLocalUserGroups!Name & "SynchTables")
110     RSSynchTable.CursorLocation = adUseClient

112     Set RSSynchTable = m_frmOASISProgress.OpenHttpCommsRS(sString, True)

114     If RSSynchTable Is Nothing Then
116         MsgBox "Server table [" & RSLocalUserGroups!Name & "SynchTables] does not exist!", vbCritical, "Server database in error"
            Exit Sub
        End If
        
118     Set dxDBGrid1.DataSource = RSSynchTable '.clone
120     dxDBGrid1.Columns.RetrieveFields
122     Set RSTableToSynch = New ADODB.Recordset

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSynchLayerWizard.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

        'RSSynchTable.Requery
100     Set RSSynchTable = Nothing
102     Set RSLocalUserGroups = Nothing
104     Set RSTableToSynch = Nothing
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSynchLayerWizard.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

