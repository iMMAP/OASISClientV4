VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OASIS Administrator Toolkit Login"
   ClientHeight    =   1920
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5595
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":6852
   ScaleHeight     =   1134.399
   ScaleMode       =   0  'User
   ScaleWidth      =   5253.402
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOasisServer 
      Height          =   345
      Left            =   3120
      TabIndex        =   7
      Top             =   90
      Width           =   2325
   End
   Begin VB.CheckBox chkChangePassword 
      BackColor       =   &H0000C000&
      Caption         =   "Change Password"
      Height          =   375
      Left            =   1890
      TabIndex        =   6
      Top             =   1350
      Width           =   1005
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   3105
      TabIndex        =   1
      Text            =   "oasisadmin"
      Top             =   495
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FF80&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3105
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1380
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1380
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3105
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   900
      Width           =   2325
   End
   Begin VB.Label lblOASISSERVER 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "OASIS SERVER:"
      Height          =   195
      Index           =   2
      Left            =   1890
      TabIndex        =   8
      Top             =   135
      Width           =   1230
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C000&
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   1905
      TabIndex        =   0
      Top             =   510
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H0000C000&
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   1905
      TabIndex        =   2
      Top             =   900
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
        'set the global var to false
        'to denote a failed login
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     LoginSucceeded = False
102     Me.Hide
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmLogin.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOK_Click()
        'check for correct password
        '<EhHeader>
        On Error GoTo cmdOK_Click_Err
        '</EhHeader>
100     If txtPassword.Text = "oasisadmin123" And txtUserName.Text = "oasisadmin" Then

102         If chkChangePassword.Value = vbChecked Then
104             MsgBox "You don't have the correct permissions to change the administration password.", vbInformation, "OASIS administrator toolkit"
            End If
        
106         LoginSucceeded = True
108         Me.Hide
        Else
110         MsgBox "Invalid Password, try again!", , "Login"
112         txtPassword.SetFocus
            'SendKeys "{Home}+{End}"
        End If
        '<EhFooter>
        Exit Sub

cmdOK_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmLogin.cmdOK_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     lblLabels(0).BackColor = RGB(164, 192, 80)
102     lblLabels(1).BackColor = RGB(164, 192, 80)
104     lblOASISServer(2).BackColor = RGB(164, 192, 80)
106     chkChangePassword.BackColor = RGB(164, 192, 80)
108     cmdCancel.BackColor = RGB(164, 192, 80)
110     cmdOK.BackColor = RGB(164, 192, 80)
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmLogin.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
