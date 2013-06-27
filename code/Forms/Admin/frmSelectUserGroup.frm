VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmSelectUserGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select User Group"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6495
   Icon            =   "frmSelectUserGroup.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6495
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5340
      TabIndex        =   2
      Top             =   2760
      Width           =   1065
   End
   Begin VB.CommandButton cmdGroupSelection 
      Caption         =   "OK"
      Height          =   315
      Left            =   4230
      TabIndex        =   0
      Top             =   2760
      Width           =   1065
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxUG 
      Height          =   2565
      Left            =   1860
      OleObjectBlob   =   "frmSelectUserGroup.frx":6852
      TabIndex        =   1
      Top             =   90
      Width           =   4575
   End
End
Attribute VB_Name = "frmSelectUserGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
    Me.Tag = False
End Sub

Private Sub cmdGroupSelection_Click()
        '<EhHeader>
        On Error GoTo cmdGroupSelection_Click_Err

        '</EhHeader>
        'If dxUG.Dataset.Bof Or dxUG.Dataset.EOF Then
            'dxUG.Dataset.First
        'End If

100     Me.Hide
        Me.Tag = True

        '<EhFooter>
        Exit Sub

cmdGroupSelection_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmSelectUserGroup.cmdGroupSelection_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxUG_OnDblClick()
        '<EhHeader>
        On Error GoTo dxUG_OnDblClick_Err
        '</EhHeader>
100     Call cmdGroupSelection_Click
        '<EhFooter>
        Exit Sub

dxUG_OnDblClick_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmSelectUserGroup.dxUG_OnDblClick " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        'If dxUG.Dataset.Bof Or dxUG.Dataset.EOF Then
            'dxUG.Dataset.First
        'End If
        Me.Tag = False
        Me.Picture = g_PictureDialogSmall
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmSelectUserGroup.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Tag = False
End Sub
