VERSION 5.00
Begin VB.Form frmTableChooser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose OASIS Table"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3390
   Icon            =   "frmTableChooser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3405
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   5100
      TabIndex        =   0
      Top             =   1830
      Width           =   4335
   End
End
Attribute VB_Name = "frmTableChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
End Sub
