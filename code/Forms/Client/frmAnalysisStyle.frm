VERSION 5.00
Begin VB.Form frmAnalysisStyle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spatial Analysis Style"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4245
   Icon            =   "frmAnalysisStyle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraPoint 
      Caption         =   "Point"
      Height          =   1320
      Left            =   45
      TabIndex        =   2
      Top             =   3195
      Width           =   4155
   End
   Begin VB.Frame FraLine 
      Caption         =   "Line:"
      Height          =   1590
      Left            =   45
      TabIndex        =   1
      Top             =   1620
      Width           =   4155
   End
   Begin VB.Frame FraRegion 
      Caption         =   "Region:"
      Height          =   1590
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      Begin OASISClient.ColorPicker ColorPicker1 
         Height          =   330
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Width           =   3885
         _extentx        =   6853
         _extenty        =   582
      End
   End
End
Attribute VB_Name = "frmAnalysisStyle"
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
