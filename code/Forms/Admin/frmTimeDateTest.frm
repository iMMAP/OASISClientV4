VERSION 5.00
Begin VB.Form frmTimeDateTest 
   BackColor       =   &H0050C0A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Server Time Date Test"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4635
   Icon            =   "frmTimeDateTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4635
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   855
      TabIndex        =   4
      Top             =   450
      Width           =   3750
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   900
      TabIndex        =   1
      Text            =   "http://www.immap.org/"
      Top             =   0
      Width           =   3705
   End
   Begin VB.CommandButton cmdCheckServer 
      Caption         =   "Check Server"
      Height          =   285
      Left            =   3375
      TabIndex        =   0
      Top             =   810
      Width           =   1230
   End
   Begin VB.Label lblServerTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Time/Date:"
      Height          =   555
      Left            =   0
      TabIndex        =   3
      Top             =   315
      Width           =   825
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "URL:"
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   510
   End
End
Attribute VB_Name = "frmTimeDateTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheckServer_Click()
    
    Dim sSQL As String
    Dim sReturnValue As String
    
    txtTime.Text = ""

    If Right$(txtURL.Text, 1) <> "/" Then
        txtURL.Text = txtURL.Text & "/"
    End If

    sSQL = txtURL.Text & "oasis.asp?servtime=1"
    sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(sSQL, True)

    txtTime.Text = sReturnValue

End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     txtURL.Text = WebSite 'GetSetting(App.EXEName, "Settings", "WebServerDateTest", "http://www.immap.org/")
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmTimeDateTest.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     SaveSetting App.EXEName, "Settings", "WebServerDateTest", txtURL.Text
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmTimeDateTest.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
