VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmDSDefinitions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create dataset definitions"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11175
   Icon            =   "frmDSDefinitions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraDatasetDefinition 
      Caption         =   "Dataset Definition"
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11130
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   8055
         TabIndex        =   3
         Top             =   7110
         Width           =   1320
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
         Height          =   375
         Left            =   9585
         TabIndex        =   2
         Top             =   7110
         Width           =   1365
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxAttrSynchDSDefinitions 
         Height          =   6720
         Left            =   135
         OleObjectBlob   =   "frmDSDefinitions.frx":6852
         TabIndex        =   1
         Top             =   270
         Width           =   10905
      End
   End
End
Attribute VB_Name = "frmDSDefinitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub hj()
    'SetAttr
End Sub
