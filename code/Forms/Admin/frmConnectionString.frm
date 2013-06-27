VERSION 5.00
Begin VB.Form frmConnectionString 
   BackColor       =   &H0050C0A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Connection String Editor"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7800
   Icon            =   "frmConnectionString.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLookupConnection 
      Caption         =   "Lookup Connection String"
      Height          =   510
      Left            =   630
      TabIndex        =   6
      Top             =   855
      Width           =   2310
   End
   Begin VB.CommandButton cmdUpdateServer 
      Caption         =   "Update Server Connection String"
      Height          =   510
      Left            =   3015
      TabIndex        =   5
      Top             =   855
      Width           =   2310
   End
   Begin VB.CommandButton cmdCheckServer 
      Caption         =   "Check Server Connection String"
      Height          =   510
      Left            =   5445
      TabIndex        =   2
      Top             =   855
      Width           =   2310
   End
   Begin VB.TextBox txtHttpWww 
      Enabled         =   0   'False
      Height          =   285
      Left            =   900
      TabIndex        =   1
      Text            =   "http://www.immap.org/"
      Top             =   0
      Width           =   6855
   End
   Begin VB.TextBox txtConnectionString 
      Height          =   285
      Left            =   855
      TabIndex        =   0
      Top             =   450
      Width           =   6900
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "URL:"
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   510
   End
   Begin VB.Label lblServerConnection 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Connection String:"
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   315
      Width           =   870
   End
End
Attribute VB_Name = "frmConnectionString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheckServer_Click()
        '<EhHeader>
        On Error GoTo cmdCheckServer_Click_Err
        '</EhHeader>
    
        Dim sSQL As String
        Dim sReturnValue As String
    
100     txtConnectionString.Text = ""

102     sSQL = txtHttpWww.Text & "oasis.asp?dns="
104     sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(sSQL, True)

106     txtConnectionString.Text = sReturnValue

        '<EhFooter>
        Exit Sub

cmdCheckServer_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConnectionString.cmdCheckServer_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdLookupConnection_Click()
        '<EhHeader>
        On Error GoTo cmdLookupConnection_Click_Err
        '</EhHeader>
100     ShellExecute Me.hwnd, vbNullString, "http://www.connectionstrings.com", vbNullString, vbNullString, 1
        '<EhFooter>
        Exit Sub

cmdLookupConnection_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConnectionString.cmdLookupConnection_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdUpdateServer_Click()
        '<EhHeader>
        On Error GoTo cmdUpdateServer_Click_Err
        '</EhHeader>
        
        Dim sSQL As String
        Dim sReturnValue As String
        
100     If MsgBox("Are you sure you wanna change the connection string?  This could be problematic!!!", vbYesNo, "BE CAREFUL!") = vbYes Then
        
102         If MsgBox("Is this connection string correct?" & Chr(13) & Chr(13) & Me.txtConnectionString, vbYesNo, "BE CAREFUL!") = vbYes Then
                
106             sSQL = txtHttpWww.Text & "oasis.asp?dns=" & CheckEncrypt(txtConnectionString.Text)
                txtConnectionString.Text = ""
108             sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(sSQL, True)
                txtConnectionString.Text = ""
110             txtConnectionString.Text = sReturnValue

            End If
        End If

        '<EhFooter>
        Exit Sub

cmdUpdateServer_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConnectionString.cmdUpdateServer_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     txtHttpWww.Text = WebSite 'etSetting(App.EXEName, "Settings", "WebServerConString", "http://www.immap.org/")
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConnectionString.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     SaveSetting App.EXEName, "Settings", "WebServerConString", txtHttpWww.Text
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConnectionString.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

