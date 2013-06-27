VERSION 5.00
Begin VB.Form frmZoomToSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1920
   Icon            =   "frmZoomToSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   1920
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSaveMarkers 
      Caption         =   "Save Markers on Exit"
      Height          =   555
      Left            =   30
      TabIndex        =   1
      Top             =   450
      Width           =   1935
   End
   Begin VB.CheckBox chkUseMultiple 
      Caption         =   "Use Multiple Markers"
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1875
   End
End
Attribute VB_Name = "frmZoomToSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSaveMarkers_Click()
    g_ZoomToSettings.SaveOnExit = IIf(chkSaveMarkers.Value = vbChecked, True, False)
End Sub

Private Sub chkUseMultiple_Click()
    DebugPrint ""
    g_ZoomToSettings.UseMultiple = IIf(chkUseMultiple.Value = vbChecked, True, False)
End Sub

Private Sub Form_Load()
    chkUseMultiple.Value = IIf(g_ZoomToSettings.UseMultiple, vbChecked, vbUnchecked)
    chkSaveMarkers.Value = IIf(g_ZoomToSettings.SaveOnExit, vbChecked, vbUnchecked)
End Sub
