VERSION 5.00
Begin VB.Form frmEULA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS EULA (End User License Agreement)"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7350
   Icon            =   "frmEULA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIDisagree 
      Caption         =   "I Disagree"
      Height          =   330
      Left            =   4770
      TabIndex        =   2
      Top             =   5940
      Width           =   1230
   End
   Begin VB.CommandButton cmdIAgree 
      Caption         =   "I Agree"
      Height          =   330
      Left            =   6075
      TabIndex        =   1
      Top             =   5940
      Width           =   1230
   End
   Begin VB.TextBox txtLicense 
      Height          =   5820
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   45
      Width           =   7260
   End
End
Attribute VB_Name = "frmEULA"
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
