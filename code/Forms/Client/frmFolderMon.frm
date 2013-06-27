VERSION 5.00
Begin VB.Form frmFoldermon 
   Caption         =   "OASIS Folder Monitor test"
   ClientHeight    =   510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   Icon            =   "frmFolderMon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   420
      Left            =   3930
      TabIndex        =   2
      Top             =   60
      Width           =   1350
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop Monitoring"
      Height          =   420
      Left            =   1530
      TabIndex        =   1
      Top             =   60
      Width           =   1350
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start &Monitoring"
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1350
   End
End
Attribute VB_Name = "frmFoldermon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OASISFolderMonitorImporter As New clFolderMonitorImporter

Private Sub cmdExit_Click()
    cmdStop_Click
    End
End Sub

Private Sub cmdStart_Click()
    cmdStart.Enabled = False
    OASISFolderMonitorImporter.CommenceMonitoring "C:\Users\IMMAP\Documents\iMMAP - OASIS\OASIS client\data\sync\import", "C:\Users\IMMAP\Documents\iMMAP - OASIS\OASIS client\data\db\OASISclient.mdb"
    cmdStart.Enabled = True
End Sub

Private Sub cmdStop_Click()
    OASISFolderMonitorImporter.StopMonitoring
End Sub

Private Sub Form_Load()
    Set OASISFolderMonitorImporter = New clFolderMonitorImporter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OASISFolderMonitorImporter.StopMonitoring
End Sub

