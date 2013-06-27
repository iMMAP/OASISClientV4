VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmUserAccounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Accounts"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7560
   FillColor       =   &H00C0C0C0&
   Icon            =   "frmUserAccounts.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   4170
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6660
      TabIndex        =   11
      Top             =   4170
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0050C0A4&
      Caption         =   "Filter by: "
      Height          =   1455
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   3615
      Begin VB.ComboBox cmbUserGroups 
         DataField       =   "Description"
         DataSource      =   "dxDBGrid"
         Height          =   315
         ItemData        =   "frmUserAccounts.frx":6852
         Left            =   240
         List            =   "frmUserAccounts.frx":6854
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtFilter 
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   3015
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "None"
         Height          =   375
         Index           =   4
         Left            =   7920
         TabIndex        =   8
         Top             =   600
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optFilter 
         BackColor       =   &H0050C0A4&
         Caption         =   "Last Name"
         Height          =   375
         Index           =   3
         Left            =   2280
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optFilter 
         BackColor       =   &H0050C0A4&
         Caption         =   "First Name"
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optFilter 
         BackColor       =   &H0050C0A4&
         Caption         =   "User Group"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optFilter 
         BackColor       =   &H0050C0A4&
         Caption         =   "Use Account Name"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   4170
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Index           =   1
      Left            =   4860
      TabIndex        =   0
      Top             =   4170
      Width           =   855
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxUsers 
      Height          =   2415
      Left            =   0
      OleObjectBlob   =   "frmUserAccounts.frx":6856
      TabIndex        =   2
      Top             =   1680
      Width           =   7575
   End
End
Attribute VB_Name = "frmUserAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSUserAccounts As New ADODB.Recordset

Public Sub setUserAccountsRS(ByVal PassedRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo setUserAccountsRS_Err
        '</EhHeader>
100     Set RSUserAccounts = PassedRS
        '<EhFooter>
        Exit Sub

setUserAccountsRS_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmUserAccounts.setUserAccountsRS " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function Wantstosaveuseraccounts() As Boolean
        '<EhHeader>
        On Error GoTo Wantstosaveuseraccounts_Err
        '</EhHeader>

        Dim sNewUAccs() As String
        Dim sDelUAccs() As String
        Dim bReturnVal As Integer
        Dim bAbortFlag As Boolean
        Dim boolReturnValue As Boolean
        
        Wantstosaveuseraccounts = True
100     bAbortFlag = False
    
102     If Not RSUserAccounts.State = adStateClosed Then
        
104         RSUserAccounts.Filter = adFilterPendingRecords

106         If Not RSUserAccounts.EOF Or Not RSUserAccounts.Bof Then
            
108             ReDim sNewUAccs(0)
110             ReDim sDelUAccs(0)
112             RSUserAccounts.MoveFirst
114             bReturnVal = MsgBox("Do you wish to save your changes?", vbYesNo, "Confirm Save")

116             If bReturnVal = vbYes Then
                
118                 Do While Not bAbortFlag And Not RSUserAccounts.EOF
                  
122                     Select Case RSUserAccounts.EditMode
            
                            Case adEditAdd
                                
                                If Not IsNull(RSUserAccounts.fields("UserGroupID").Value) Then
124                                 sNewUAccs(UBound(sNewUAccs)) = RSUserAccounts.fields.Item("user").Value
126                                 RSUserAccounts.fields.Item("sGUID").Value = GUIDGen()
128                                 ReDim Preserve sNewUAccs(UBound(sNewUAccs) + 1)
                                Else
                                    bAbortFlag = True

                                End If
    
130                         Case adEditDelete
132                             sDelUAccs(UBound(sDelUAccs)) = RSUserAccounts.fields.Item("user").OriginalValue
134                             ReDim Preserve sDelUAccs(UBound(sDelUAccs) + 1)
                        
                        End Select
                    
136                     RSUserAccounts.MoveNext

                    Loop
                    
140                 If bAbortFlag Then
                    
142                     MsgBox "All users MUST have a usergroup assigned to them. Save aborted", vbCritical, "No user group specified"
                        Wantstosaveuseraccounts = False
                    Else

144                     If Not RSUserAccounts.EOF Or Not RSUserAccounts.Bof Then RSUserAccounts.MoveFirst
146                     boolReturnValue = m_frmOASISProgress.SaveHttpCommsRS(RSUserAccounts, WebSite & "Oasis.asp", True)
    
148                     If boolReturnValue Then
150                         MsgBox "User accounts successfully updated on server"
152                         SetStatus "User accounts successfully updated on server"
                        Else
154                         MsgBox "User accounts unsuccessfully updated on server", vbCritical
156                         SetStatus "User accounts unsuccessfully updated on server"
                        End If

158
                    End If

                Else
                    Wantstosaveuseraccounts = False
                End If
            
            End If
        
160         RSUserAccounts.Filter = ""

162         If Not RSUserAccounts.Bof Or Not RSUserAccounts.EOF Then RSUserAccounts.MoveFirst
        
        End If

        '<EhFooter>
        Exit Function

Wantstosaveuseraccounts_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmUserAccounts.Wantstosaveuseraccounts " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmbUserGroups_Change()
        '<EhHeader>
        On Error GoTo cmbUserGroups_Change_Err
        '</EhHeader>
100     txtFilter_Change
        '<EhFooter>
        Exit Sub

cmbUserGroups_Change_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmUserAccounts.cmbUserGroups_Change " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmbUserGroups_Click()
        '<EhHeader>
        On Error GoTo cmbUserGroups_Click_Err
        '</EhHeader>
100     txtFilter_Change
        '<EhFooter>
        Exit Sub

cmbUserGroups_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmUserAccounts.cmbUserGroups_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
        'RSUserAccounts.Requery
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmUserAccounts.cmdCancel_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdDelete_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo cmdDelete_Click_Err

        '</EhHeader>
100     If Not dxUsers.Dataset.IsEmpty Then

102         If MsgBox("Are you sure you want to delete the selected user account?", vbYesNo, "Confirm Deletion") = vbYes Then
104             dxUsers.Dataset.Delete
106             dxUsers.Dataset.First
            End If

        Else
108         MsgBox "There is no selected user account!", vbExclamation, "Error"
        End If

        '<EhFooter>
        Exit Sub

cmdDelete_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmUserAccounts.cmdDelete_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdNew_Click()
        '<EhHeader>
        On Error GoTo cmdNew_Click_Err
        '</EhHeader>

        Dim sString As String
        Dim bTryAgain As Boolean
    
        Do
100         bTryAgain = False
102         Me.optFilter(4) = True
104         txtFilter_Change
    
106         sString = InputBox("Please enter a new User Account Name", "New User Account")

108         If Not IsNull(sString) And Not sString = "" Then

110             dxUsers.Filter.AddFirst 0, otEqual, sString, sString, False
112             dxUsers.Filter.Apply
            
114             If dxUsers.Dataset.EOF Then
            
116                 Me.dxUsers.Dataset.Append
118                 Me.dxUsers.Dataset.FieldValues("user") = sString
120                 Me.dxUsers.Dataset.FieldValues("pwd") = "newpassword"
122                 Me.dxUsers.Dataset.FieldValues("Fname") = "newfname"
124                 Me.dxUsers.Dataset.FieldValues("Lname") = "newlname"
126                 Me.dxUsers.Dataset.FieldValues("SettingUrl") = "newsettingurl"
128                 Me.dxUsers.Dataset.Post
130                 dxUsers.Dataset.Refresh
            
                Else
            
132                 MsgBox "This user account already exists!", vbCritical, "Error"
134                 bTryAgain = True
                
                End If
            
136             dxUsers.Filter.Clear

            End If

138     Loop Until bTryAgain = False

        '<EhFooter>
        Exit Sub

cmdNew_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmUserAccounts.cmdNew_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOK_Click()
        '<EhHeader>
        On Error GoTo cmdOK_Click_Err
        '</EhHeader>

        If dxUsers.Dataset.State = 2 Then
            dxUsers.Dataset.Post
        End If
        
104     If Wantstosaveuseraccounts Then
            Me.Visible = False
            Unload Me
        End If

        '<EhFooter>
        Exit Sub

cmdOK_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmUserAccounts.cmdOK_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     txtFilter_Change
102     Me.Picture = g_PictureDialogSmall
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmUserAccounts.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmDatabaseConnect.cmdConnect_Click
End Sub

Private Sub optFilter_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo optFilter_Click_Err
        '</EhHeader>
100     txtFilter_Change
        '<EhFooter>
        Exit Sub

optFilter_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmUserAccounts.optFilter_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtFilter_Change()
        '<EhHeader>
        On Error GoTo txtFilter_Change_Err
        '</EhHeader>

100     If Me.optFilter(1) Then
102         Me.dxUsers.Dataset.Filtered = True
104         txtFilter.BackColor = &HC0C0C0
106         txtFilter.Visible = False
108         cmbUserGroups.Visible = True
110     ElseIf Me.optFilter(4) Then
112         Me.dxUsers.Dataset.Filtered = False
114         Me.txtFilter.Enabled = False
116         txtFilter.BackColor = &HC0C0C0
118         txtFilter.Visible = True
120         cmbUserGroups.Visible = False
        Else
122         Me.dxUsers.Dataset.Filtered = False
124         Me.txtFilter.Enabled = True
126         txtFilter.BackColor = vbWhite
128         txtFilter.Visible = True
130         cmbUserGroups.Visible = False
        End If
    
132     If Me.optFilter(0) And Not txtFilter = "" Then

134         Me.dxUsers.Dataset.Filter = "user LIKE '" & Me.txtFilter & "*'"
136         Me.dxUsers.Dataset.Filtered = True
138         txtFilter.BackColor = &H80FF80

140     ElseIf Me.optFilter(1) And Not cmbUserGroups.ListIndex < 0 Then

142         Me.dxUsers.Dataset.Filter = "UserGroupID = " & cmbUserGroups.ItemData(cmbUserGroups.ListIndex)
144         Me.dxUsers.Dataset.Filtered = True
146         txtFilter.BackColor = &H80FF80

148     ElseIf Me.optFilter(2) And Not txtFilter = "" Then

150         Me.dxUsers.Dataset.Filter = "Fname LIKE '" & Me.txtFilter & "*'"
152         Me.dxUsers.Dataset.Filtered = True
154         txtFilter.BackColor = &H80FF80

156     ElseIf Me.optFilter(3) And Not txtFilter = "" Then

158         Me.dxUsers.Dataset.Filter = "Lname LIKE '" & Me.txtFilter & "*'"
160         Me.dxUsers.Dataset.Filtered = True
162         txtFilter.BackColor = &H80FF80

        End If

        '<EhFooter>
        Exit Sub

txtFilter_Change_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmUserAccounts.txtFilter_Change " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
