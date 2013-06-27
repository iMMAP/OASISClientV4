VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmServerStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Server Status"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11760
   Icon            =   "frmServerStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11760
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCheckServer 
      Caption         =   "Check Server"
      Height          =   285
      Left            =   10440
      TabIndex        =   2
      Top             =   45
      Width           =   1230
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   630
      TabIndex        =   1
      Text            =   "http://www.immap.org/"
      Top             =   45
      Width           =   9735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8025
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   11580
      ExtentX         =   20426
      ExtentY         =   14155
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblURL 
      Caption         =   "URL:"
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   510
   End
End
Attribute VB_Name = "frmServerStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheckServer_Click()
        '<EhHeader>
        On Error GoTo cmdCheckServer_Click_Err
        '</EhHeader>
        
        If Right$(txtURL.Text, 1) <> "/" Then
            txtURL.Text = txtURL.Text & "/"
        End If
        
100     WebBrowser1.Navigate txtURL.Text & "oasis.asp?servstat=1"
        m_frmDebug.DebugPrint txtURL.Text & "oasis.asp?serverstat=1"
        '<EhFooter>
        Exit Sub

cmdCheckServer_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmServerStatus.cmdCheckServer_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     txtURL.Text = WebSite 'GetSetting(App.EXEName, "Settings", "WebServerStatus", "http://www.immap.org/")
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmServerStatus.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     SaveSetting App.EXEName, "Settings", "WebServerStatus", txtURL.Text

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmServerStatus.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
