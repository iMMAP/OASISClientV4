VERSION 5.00
Begin VB.Form frmAddLocToW3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add location"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12465
   Icon            =   "frmAddLocToW3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   12465
   StartUpPosition =   3  'Windows Default
   Begin OASISClient.ctrWhere ctrWhere1 
      Height          =   6900
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   12390
      _ExtentX        =   10504
      _ExtentY        =   12171
   End
End
Attribute VB_Name = "frmAddLocToW3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub Form_Load()
'    If Not g_sLanguage = "" Then
'        If Not m_Cnn.State = adStateClosed Then
'            LoadLanguage Me.Name, g_sLanguage, m_Cnn
'        End If
'    End If
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    If MsgBox("Do you really want to attach this location?", vbYesNo) = vbYes Then
'        ctrWhere1.UpdateRecord
'    Else
'        ctrWhere1.DeleteRecord
'    End If
'End Sub
