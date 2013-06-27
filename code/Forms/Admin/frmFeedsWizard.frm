VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmFeedsWizard 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS RSS Feed Editor"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7560
   Icon            =   "frmFeedsWizard.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox ComRSSGroups 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdNewCategory 
      Caption         =   "..."
      Height          =   315
      Left            =   7050
      TabIndex        =   13
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   6660
      TabIndex        =   12
      Top             =   4290
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   285
      Left            =   4860
      TabIndex        =   11
      Top             =   4290
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Add"
      Height          =   285
      Left            =   5760
      TabIndex        =   10
      Top             =   4290
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   0
      Left            =   4080
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2265
      Left            =   120
      OleObjectBlob   =   "frmFeedsWizard.frx":6852
      TabIndex        =   15
      Top             =   1920
      Width           =   7365
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0050C0A4&
      BackStyle       =   0  'Transparent
      Caption         =   "Group ID:"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   8
      Top             =   1590
      Width           =   1845
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0050C0A4&
      BackStyle       =   0  'Transparent
      Caption         =   "Feed URL:"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0050C0A4&
      BackStyle       =   0  'Transparent
      Caption         =   "Feed Name:"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0050C0A4&
      BackStyle       =   0  'Transparent
      Caption         =   "Feed Image URL:"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0050C0A4&
      BackStyle       =   0  'Transparent
      Caption         =   "Feed Description:"
      Height          =   285
      Index           =   0
      Left            =   2130
      TabIndex        =   0
      Top             =   1200
      Width           =   1875
   End
End
Attribute VB_Name = "frmFeedsWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RSLocalUserGroups As ADODB.Recordset
Private RSFeeds As ADODB.Recordset
Private RsFeedGroups As ADODB.Recordset
Private m_bIsInProgress As Boolean
Private WithEvents m_frmAddFeedGroups As frmAddFeedGroups
Attribute m_frmAddFeedGroups.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    CheckEdit (False)
    Unload Me

End Sub

Private Sub cmdNew_Click()
        
    If ComRSSGroups.ListCount < 1 Or ComRSSGroups.ListIndex < 0 Then
        MsgBox "You have no valid chosen Feed Group. You need to add at least 1 Feed Group To continue", vbOKOnly, "OASIS Feeds Wizard..."
        Exit Sub
    End If
        
    RSFeeds.AddNew
    RSFeeds.fields("FeedDescription") = txtFields(0).Text
    RSFeeds.fields("FeedImageURL") = txtFields(1).Text
    RSFeeds.fields("FeedName") = txtFields(2).Text
    RSFeeds.fields("FeedURL") = txtFields(3).Text
    RSFeeds.fields("GroupID") = Me.ComRSSGroups.ItemData(ComRSSGroups.ListIndex)   'txtFields(4).Text

    txtFields(0).Text = ""
    txtFields(1).Text = ""
    txtFields(2).Text = ""
    txtFields(3).Text = ""
        
End Sub

Private Sub cmdNewCategory_Click()

    Set m_frmAddFeedGroups = New frmAddFeedGroups
    Call m_frmAddFeedGroups.setUserGroupsRS(RSLocalUserGroups)
    Call m_frmAddFeedGroups.setFeedGroupsRS(RsFeedGroups)
    m_frmAddFeedGroups.Show vbModeless, Me

End Sub

Private Sub ComRSSGroups_Click()

    If Not m_bIsInProgress Then
        txtFields(4).Text = ComRSSGroups.ItemData(ComRSSGroups.ListIndex)
    End If
    
End Sub

Private Sub dxDBGrid1_OnDblClick()
    
    If DeleteRecordFromRSAndSave(RSFeeds, "SettingValue5", RSLocalUserGroups.fields("Name").Value) Then Unload Me

End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
        Dim sString As String
        Dim i As Integer

100     Me.Picture = g_PictureDialogSmall
102     m_bIsInProgress = True

104     sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & RSLocalUserGroups!Name & "Feeds")
106     Set RSFeeds = m_frmOASISProgress.OpenHttpCommsRS(sString, True)

108     If RSFeeds.State = adStateClosed Then
110         MsgBox "Something Shifty went on @ the server....!"
            Exit Sub
        End If

112     Call m_frmAddFeedGroups_PopulateFeedsCombo
114     Set dxDBGrid1.DataSource = RSFeeds
116     dxDBGrid1.KeyField = "FeedID"
118     dxDBGrid1.Columns.RetrieveFields

120     dxDBGrid1.Columns(0).Visible = False
122     dxDBGrid1.Columns(1).Visible = False
124     dxDBGrid1.Columns(2).Visible = False
126     dxDBGrid1.Columns(3).Visible = True
128     dxDBGrid1.Columns(4).Visible = True
130     dxDBGrid1.Columns(5).Visible = True
132     dxDBGrid1.Columns(6).Visible = True
134     dxDBGrid1.Columns(7).Visible = False
136     dxDBGrid1.Columns(8).Visible = False
138     dxDBGrid1.Columns(9).Visible = False

140     dxDBGrid1.Columns(3).Width = dxDBGrid1.Width / 4
142     dxDBGrid1.Columns(4).Width = dxDBGrid1.Width / 4
144     dxDBGrid1.Columns(5).Width = dxDBGrid1.Width / 4
146     dxDBGrid1.Columns(6).Width = dxDBGrid1.Width / 4

148     m_bIsInProgress = False
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmFeedsWizard.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub m_frmAddFeedGroups_PopulateFeedsCombo()

    Dim sString As String

    sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & RSLocalUserGroups!Name & "FeedGroups")
    Set RsFeedGroups = m_frmOASISProgress.OpenHttpCommsRS(sString, True)

    If RsFeedGroups.State = adStateClosed Then
        MsgBox "Something Shifty went on @ the server....!"
        Exit Sub
    End If

    ComRSSGroups.Clear
    
    If Not RsFeedGroups.EOF Or Not RsFeedGroups.Bof Then
    
        RsFeedGroups.MoveFirst
        
        Do Until RsFeedGroups.EOF
            ComRSSGroups.AddItem RsFeedGroups.fields("GroupText")
            ComRSSGroups.ItemData(ComRSSGroups.ListCount - 1) = RsFeedGroups.fields("GroupId")
            RsFeedGroups.MoveNext
        Loop
        
        If Not ComRSSGroups.ListCount = 0 Then ComRSSGroups.ListIndex = 0

    End If

End Sub

Public Sub setUserGroupsRS(ByVal PassedRS As ADODB.Recordset)
    Set RSLocalUserGroups = PassedRS
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RSFeeds = Nothing
    Set RSLocalUserGroups = Nothing
    Set RsFeedGroups = Nothing
End Sub

Private Function CheckEdit(bPrompt As Boolean) As Boolean
    
    Dim msgBoxResult As VbMsgBoxResult
    Dim bReturnValue As Boolean

    CheckEdit = False

    With RSFeeds
        .Filter = adFilterPendingRecords
    
        If Not .Bof And Not .EOF Then

            If bPrompt Then
                msgBoxResult = MsgBox("Do you wish to save your changes?", vbYesNoCancel, "Confirm Save")
            Else
                msgBoxResult = vbYes
            End If

            Select Case msgBoxResult
                
                Case vbYes
                
                    bReturnValue = m_frmOASISProgress.SaveHttpCommsRS(RSFeeds, WebSite & "Oasis.asp", True)
                    
                    If bReturnValue Then
                        IncrementProfileSettingVersion WebSite, "SettingValue5", RSLocalUserGroups.fields("Name").Value
                        MsgBox "Data saved to server"
                    Else
                        MsgBox "Saving to server failed!"
                    End If

                    CheckEdit = True

                Case vbNo
                
                    CheckEdit = True
            End Select
                
        Else
            CheckEdit = True
        End If

    End With

End Function
