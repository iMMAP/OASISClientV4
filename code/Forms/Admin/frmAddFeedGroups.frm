VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmAddFeedGroups 
   BackColor       =   &H0050C0A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Feed Groups"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3360
   FillColor       =   &H00C0FFC0&
   Icon            =   "frmAddFeedGroups.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2535
      Left            =   0
      OleObjectBlob   =   "frmAddFeedGroups.frx":6852
      TabIndex        =   5
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   3180
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   285
      Left            =   2580
      TabIndex        =   2
      Top             =   3180
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   2850
      Width           =   3285
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   285
      Left            =   1830
      TabIndex        =   0
      Top             =   3180
      Width           =   735
   End
   Begin VB.Label lblName 
      BackColor       =   &H0050C0A4&
      Caption         =   "Name:"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   2610
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddFeedGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsFeedGroups As ADODB.Recordset
Private RSLocalUserGroups As ADODB.Recordset
Public Event PopulateFeedsCombo()

Private Sub cmdDelete_Click()

    If MsgBox("Do you want to delete this category?", vbYesNo, "OASIS ADmin Tool") = vbYes Then
        RsFeedGroups.Delete
    End If

End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Function CheckEdit() As Boolean

    Dim bReturnValue As Boolean

    With RsFeedGroups
        .Filter = adFilterPendingRecords

        If Not .Bof And Not .EOF Then

            Select Case MsgBox("Do you wish to save your changes?", vbYesNoCancel, "Confirm Save")

                Case vbYes
                    
                    bReturnValue = m_frmOASISProgress.SaveHttpCommsRS(RsFeedGroups, WebSite & "Oasis.asp", True)

                    If bReturnValue Then
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

Private Sub cmdNew_Click()

    RsFeedGroups.AddNew
    RsFeedGroups.fields("GroupText").Value = txtName
    RsFeedGroups.fields("CustomGroup").Value = True
    Set dxDBGrid1.DataSource = RsFeedGroups
    dxDBGrid1.Columns.RetrieveFields

End Sub

Public Sub setFeedGroupsRS(ByRef RsFeedGroupsPassed As ADODB.Recordset)
        
    Set dxDBGrid1.DataSource = Nothing
    Set RsFeedGroups = RsFeedGroupsPassed
    Set dxDBGrid1.DataSource = RsFeedGroups
    dxDBGrid1.Columns.RetrieveFields

End Sub

Public Sub setUserGroupsRS(ByVal PassedRS As ADODB.Recordset)
    Set RSLocalUserGroups = PassedRS
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not CheckEdit Then Cancel = 1
    Set RsFeedGroups = Nothing
    Set RSLocalUserGroups = Nothing
    RaiseEvent PopulateFeedsCombo
    
End Sub

