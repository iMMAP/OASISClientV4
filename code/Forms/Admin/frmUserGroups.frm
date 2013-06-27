VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmUserGroups 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Groups"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7545
   Icon            =   "frmUserGroups.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5970
      TabIndex        =   4
      Top             =   3990
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6750
      TabIndex        =   3
      Top             =   3990
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5190
      TabIndex        =   2
      Top             =   3990
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   4410
      Picture         =   "frmUserGroups.frx":6852
      TabIndex        =   1
      Top             =   3990
      Width           =   735
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxUG 
      Height          =   2205
      Left            =   0
      OleObjectBlob   =   "frmUserGroups.frx":48744
      TabIndex        =   0
      Top             =   1680
      Width           =   7575
   End
End
Attribute VB_Name = "frmUserGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSUserGroups As New ADODB.Recordset
Dim RSLocalUserAccounts As New ADODB.Recordset
Dim sFeedback As String

Public Sub setUserGroupsRS(ByVal PassedRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo setUserGroupsRS_Err
        '</EhHeader>
100     Set RSUserGroups = PassedRS
102     Set dxUG.DataSource = RSUserGroups
104     dxUG.Columns.RetrieveFields
        '<EhFooter>
        Exit Sub

setUserGroupsRS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.setUserGroupsRS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub setUserAccountsRS(ByVal PassedRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo setUserAccountsRS_Err
        '</EhHeader>
100     Set RSLocalUserAccounts = PassedRS
        '<EhFooter>
        Exit Sub

setUserAccountsRS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.setUserAccountsRS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub doProgress(bStart As Boolean, _
                       Optional sText As String, _
                       Optional lProc As Long)
        '<EhHeader>
        On Error GoTo doProgress_Err
        '</EhHeader>

        On Error Resume Next
    
100     SetStatus "Begin Progress..."
    
102     If bStart Then

104         frmProgress.Show vbModeless, Me
106         With frmProgress
108             .Timer1.Enabled = bStart
110             .dxProgressBar.Visible = bStart
            End With
    
        Else
112         frmProgress.Hide
114         Unload frmProgress
        End If
        
        '<EhFooter>
        Exit Sub

doProgress_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.doProgress " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CheckIfNameOfGroupIsOK(sGroupName As String) As Boolean
        '<EhHeader>
        On Error GoTo CheckIfNameOfGroupIsOK_Err
        '</EhHeader>

        Dim RSTables As ADODB.Recordset
100     Set RSTables = m_frmOASISProgress.OpenHttpCommsRS(WebSite & "/oasis.asp?gettables", True)

102     If Not RSTables Is Nothing Then
    
104         RSTables.Find "TABLE_NAME LIKE '" & sGroupName & "%'"
        
106         If RSTables.EOF Then
108             CheckIfNameOfGroupIsOK = True
            Else
110             CheckIfNameOfGroupIsOK = False
112             sFeedback = "Either there is another group (or table) named or beginning with '" & sGroupName & "'" & Chr(13) & Chr(13) & "User Group vaildation has failed"
                Exit Function
            End If
    
        Else
    
114         sFeedback = "Validation of groupname failed due to server communication error"
116         CheckIfNameOfGroupIsOK = False
            Exit Function
        End If
        
118     If SafeMoveFirst(RSUserGroups) Then
        
120         Do Until RSUserGroups.EOF
            
122             If sGroupName Like (RSUserGroups.fields("Name").Value & "*") Then
124                 CheckIfNameOfGroupIsOK = False
126                 sFeedback = "A user group called " & RSUserGroups.fields("Name").Value & " exists." & Chr(13) & "This is in conflict with user group naming rules."
                End If
            
128             RSUserGroups.MoveNext
            Loop
            
130         SafeMoveFirst RSUserGroups
        
        End If
    
132     Set RSTables = Nothing

        '<EhFooter>
        Exit Function

CheckIfNameOfGroupIsOK_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.CheckIfNameOfGroupIsOK " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function Wantstosaveusergroups()
        '<EhHeader>
        On Error GoTo Wantstosaveusergroups_Err
        '</EhHeader>

        Dim sNewUGroups() As String
        Dim sDelUGroups() As String
        Dim bReturnVal As Integer
        Dim boolReturnValue As Boolean
        Dim i As Integer
        
100     Wantstosaveusergroups = True
    
102     If Not RSUserGroups.State = adStateClosed Then
        
104         RSUserGroups.Filter = adFilterPendingRecords

106         If Not RSUserGroups.EOF And Not RSUserGroups.Bof Then
            
108             ReDim sNewUGroups(0)
110             ReDim sDelUGroups(0)
112             RSUserGroups.MoveFirst
114             bReturnVal = MsgBox("Do you wish to save your changes?", vbYesNo, "Confirm Save")

116             If bReturnVal = vbYes Then

118                 Do While Not RSUserGroups.EOF
                        
120                     Select Case RSUserGroups.EditMode
            
                            Case adEditAdd
122                             RSUserGroups.fields.Item("SettingTablePrefix").Value = RSUserGroups.fields.Item("Name").Value
124                             RSUserGroups.fields.Item("sGUID").Value = GUIDGen()
126                             sNewUGroups(UBound(sNewUGroups)) = RSUserGroups.fields.Item("Name").Value
128                             ReDim Preserve sNewUGroups(UBound(sNewUGroups) + 1)
    
130                         Case adEditDelete
132                             sDelUGroups(UBound(sDelUGroups)) = RSUserGroups.fields.Item("Name").OriginalValue
134                             ReDim Preserve sDelUGroups(UBound(sDelUGroups) + 1)
                        
                        End Select
                    
136                     RSUserGroups.MoveNext
                    Loop
                    
138                 If Not RSUserGroups.EOF Or Not RSUserGroups.Bof Then RSUserGroups.MoveFirst
140                 boolReturnValue = m_frmOASISProgress.SaveHttpCommsRS(RSUserGroups, WebSite & "Oasis.asp", True)
    
142                 If boolReturnValue Then
                        
144                     SetStatus "User Groups successfully updated on server..."
    
146                     If UBound(sDelUGroups) > 0 Then

148                         For i = 0 To UBound(sDelUGroups) - 1

150                             If Not DeleteUserGTable(sDelUGroups(i)) Then
152                                 SetStatus "Failed deleting table:" & sDelUGroups(i) & " on server. Will now exit..."
154                                 MsgBox "Failed deleting table:" & sDelUGroups(i) & " on server. Will now exit..."
                                    Exit Function
                                Else
156                                 SetStatus "Deleting user group: " & sDelUGroups(i)
                                End If

                            Next

                        End If
        
158                     If UBound(sNewUGroups) > 0 Then
            
160                         For i = 0 To UBound(sNewUGroups) - 1

162                             If Not CreateUserGTable(sNewUGroups(i)) Then
164                                 SetStatus "Failed creating table:" & sNewUGroups(i) & " on server. Will now exit..."
166                                 MsgBox "Failed creating table:" & sNewUGroups(i) & " on server. Will now exit..."
                                    Exit Function
                                Else
168                                 SetStatus "Creating user group: " & sNewUGroups(i)
                                End If

                            Next

                        End If
                    End If

170                 MsgBox "Data updated successfully on server"
                     
                Else
172                 Wantstosaveusergroups = False
                End If
            
            End If
        
174         RSUserGroups.Filter = ""
        
        End If
                 
        '<EhFooter>
        Exit Function

Wantstosaveusergroups_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.Wantstosaveusergroups " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function CreateUserGTable(sTname As String) As Boolean
        '<EhHeader>
        On Error GoTo CreateUserGTable_Err
        '</EhHeader>

        Dim sReturnValue As String
        Dim sString As String
        
100     CreateUserGTable = False
102     SetStatus "Creating Usergroup Tables..."
104     sString = WebSite & "Oasis.asp?createug=" & CheckEncrypt(sTname) & "&from=" & CheckEncrypt("iMMAP")
106     sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(sString, True)

108     If Not Trim$(sReturnValue) = "done" Then
110         SetStatus "Creating user tables failed (response: " & sReturnValue & ")"
        Else
112         SetStatus "Creating user tables successful (response: " & sReturnValue & ")"
114         CreateUserGTable = True
        End If

        '<EhFooter>
        Exit Function

CreateUserGTable_Err:

        '</EhFooter>
End Function

Private Function DeleteUserGTable(sTname As String) As Boolean
        '<EhHeader>
        On Error GoTo DeleteUserGTable_Err
        '</EhHeader>
        
        Dim sReturnValue As String
        Dim sString As String
    
100     DeleteUserGTable = False
102     SetStatus "Deleting Usergroup Tables..."
104     sString = WebSite & "Oasis.asp?removeug=" & CheckEncrypt(sTname)
106     sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(sString, True)

108     If Not Trim$(sReturnValue) = "done" Then
110         SetStatus "Deleting user tables failed (response: " & sReturnValue & ")"
        Else
112         SetStatus "Deleting user tables successful (response: " & sReturnValue & ")"
114         DeleteUserGTable = True
        End If

        '<EhFooter>
        Exit Function

DeleteUserGTable_Err:

        '</EhFooter>
End Function

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function DoesIfUserGroupHasUsers(sUserID As String) As Boolean
        '<EhHeader>
        On Error GoTo DoesIfUserGroupHasUsers_Err
        '</EhHeader>

100     With RSLocalUserAccounts
    
102         If .State = adStateOpen And (Not .EOF Or Not .Bof) Then .MoveFirst
104         .Find "UserGroupID = " & sUserID
    
106         If Not .EOF Then
108             DoesIfUserGroupHasUsers = True
            Else
110             DoesIfUserGroupHasUsers = False
            End If

        End With

        '<EhFooter>
        Exit Function

DoesIfUserGroupHasUsers_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.DoesIfUserGroupHasUsers " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdDelete_Click()
        '<EhHeader>
        On Error GoTo cmdDelete_Click_Err
        '</EhHeader>
    
100     If Not dxUG.Dataset.IsEmpty Then
    
102         If MsgBox("Are you sure you want to delete the selected user group?", vbYesNo, "Confirm Deletion") = vbYes Then

104             If Not DoesIfUserGroupHasUsers(RSUserGroups.fields("ID").Value) Then
106                 RSUserGroups.Delete
                Else
108                 MsgBox "This user group has user accounts assigned to it.  Delete or reassign these users accounts to another user group before you delete this usergroup", vbCritical, "Usergroup has users"
                End If

                'RSUserGroups.MoveFirst
            End If
        
        Else
110         MsgBox "There is no selected user group!", vbExclamation, "Error"
        End If
    
        '<EhFooter>
        Exit Sub

cmdDelete_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.cmdDelete_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdNew_Click()
        '<EhHeader>
        On Error GoTo cmdNew_Click_Err
        '</EhHeader>
    
        Dim bTryAgain As Boolean
    
        Do
100         bTryAgain = False
102         frmDialogWithTwoFields.Caption = "New User Group"
104         frmDialogWithTwoFields.lbl1.Caption = "Please enter a unique new user group name" & Chr(13) & "(duplicate group names are not permitted)"
106         frmDialogWithTwoFields.lbl2.Caption = Chr(13) & "Please enter a user group description"
    
108         frmDialogWithTwoFields.Show vbModal, Me
    
110         If frmDialogWithTwoFields.bClickedOK Then

112             If Not CheckIfNameOfGroupIsOK(frmDialogWithTwoFields.sText1) Then
114                 MsgBox sFeedback
                    Exit Sub
                End If
    
116             Me.dxUG.Filter.AddFirst 1, otEqual, frmDialogWithTwoFields.sText1, frmDialogWithTwoFields.sText1, False
118             Me.dxUG.Filter.Apply
        
120             If dxUG.Dataset.EOF Then
122                 Me.dxUG.Dataset.Append
124                 Me.dxUG.Dataset.FieldValues("Name") = frmDialogWithTwoFields.sText1
126                 Me.dxUG.Dataset.FieldValues("Description") = frmDialogWithTwoFields.sText2
128                 Me.dxUG.Dataset.FieldValues("SettingTablePrefix") = Me.dxUG.Dataset.FieldValues("Name")
130                 Me.dxUG.Dataset.Post
132                 dxUG.Dataset.Refresh
                Else
134                 MsgBox "This user group name already exists!", vbCritical, "Error"
136                 bTryAgain = True
                End If
        
138             Me.dxUG.Filter.Clear
    
            End If
    
140     Loop Until bTryAgain = False

        '<EhFooter>
        Exit Sub

cmdNew_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.cmdNew_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOK_Click()
        '<EhHeader>
        On Error GoTo cmdOK_Click_Err
        '</EhHeader>

100     If dxUG.Dataset.State = 2 Then dxUG.Dataset.Post
102     If Wantstosaveusergroups Then Unload Me
        '<EhFooter>
        Exit Sub

cmdOK_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.cmdOK_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     Me.Picture = g_PictureDialogSmall
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     frmDatabaseConnect.cmdConnect_Click
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmUserGroups.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
