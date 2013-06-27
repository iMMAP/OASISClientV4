VERSION 5.00
Begin VB.Form frmForcedLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmForcedLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmForcedLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event DoForcedLogin(sUserName As String, sPassword As String)
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     LoginSucceeded = False
102     Me.Hide
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmForcedLogin.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOK_Click()
        '<EhHeader>
        On Error GoTo cmdOK_Click_Err
        '</EhHeader>

        Dim sString As String
        Dim sResult As String
    
100     LoginSucceeded = False
102     RaiseEvent DoForcedLogin(txtUserName.Text, txtPassword.Text)
    
104     If LoginSucceeded Then
106         Me.Hide
        Else
108         MsgBox "Login failed!", vbCritical
        End If

        '<EhFooter>
        Exit Sub

cmdOK_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmForcedLogin.cmdOK_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
