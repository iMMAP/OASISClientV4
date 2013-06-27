VERSION 5.00
Begin VB.Form frmLOG 
   BackColor       =   &H0050C0A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Application Log"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4515
   Icon            =   "frmLOG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "Clear log"
      Height          =   330
      Left            =   3285
      TabIndex        =   1
      Top             =   3015
      Width           =   1185
   End
   Begin VB.TextBox txtLog 
      Height          =   2895
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   45
      Width           =   4425
   End
End
Attribute VB_Name = "frmLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClearLog_Click()
        '<EhHeader>
        On Error GoTo cmdClearLog_Click_Err
        '</EhHeader>
100     txtLog.Text = ""
        '<EhFooter>
        Exit Sub

cmdClearLog_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmLOG.cmdClearLog_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
