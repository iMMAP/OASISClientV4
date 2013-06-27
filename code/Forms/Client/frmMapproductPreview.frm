VERSION 5.00
Begin VB.Form frmMapproductPreview 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Map preview"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   7350
      Left            =   45
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11040
   End
End
Attribute VB_Name = "frmMapproductPreview"
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

Private Sub Form_Resize()
    Image1.Move 0, 0, Me.Width - 10, Me.Height - 10
End Sub
