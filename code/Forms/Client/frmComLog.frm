VERSION 5.00
Begin VB.Form frmComLog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "OASIS Comms Log"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Keep trace"
      Height          =   255
      Left            =   4170
      TabIndex        =   1
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox txtComms 
      Height          =   2055
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   30
      Width           =   6135
   End
End
Attribute VB_Name = "frmComLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetProp Me.hwnd, "frmComLog", ObjectPtr(Me)
End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    Me.Visible = False
'    Cancel = 1
'End Sub
