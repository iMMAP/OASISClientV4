VERSION 5.00
Begin VB.Form frmExportFormats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2040
   Icon            =   "frmExportFormats.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   2040
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraExportFormats 
      Caption         =   "Export Formats:"
      Height          =   1590
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CheckBox chkGPSExport 
         Caption         =   "GPS Export GPX"
         Height          =   225
         Left            =   135
         TabIndex        =   13
         Top             =   1320
         Width           =   1785
      End
      Begin VB.CheckBox chkMapInfoTAB 
         Caption         =   "OGIS GML"
         Height          =   330
         Left            =   135
         TabIndex        =   11
         Top             =   1020
         Width           =   1590
      End
      Begin VB.CheckBox chkAutocadDWG 
         Caption         =   "Autocad DXF"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   765
         Width           =   1545
      End
      Begin VB.CheckBox chkGoogleKML 
         Caption         =   "Google KML"
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   495
         Width           =   1635
      End
      Begin VB.CheckBox chkShapeFile 
         Caption         =   "ESRI Shape SHP"
         Height          =   240
         Left            =   135
         TabIndex        =   8
         Top             =   270
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   240
      Left            =   1050
      TabIndex        =   6
      Top             =   1605
      Width           =   960
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   240
      Left            =   60
      TabIndex        =   5
      Top             =   1605
      Width           =   960
   End
   Begin VB.Frame FraChooseExport 
      Caption         =   "Choose Export Formats:"
      Height          =   1590
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   1995
      Begin VB.CheckBox chkFormats 
         Caption         =   "OASIS Reports"
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   1320
         Width           =   1590
      End
      Begin VB.CheckBox chkFormats 
         Caption         =   "Text"
         Height          =   300
         Index           =   3
         Left            =   90
         TabIndex        =   4
         Top             =   1050
         Width           =   1590
      End
      Begin VB.CheckBox chkFormats 
         Caption         =   "HTML"
         Height          =   330
         Index           =   2
         Left            =   90
         TabIndex        =   3
         Top             =   765
         Width           =   1590
      End
      Begin VB.CheckBox chkFormats 
         Caption         =   "XML"
         Height          =   330
         Index           =   1
         Left            =   90
         TabIndex        =   2
         Top             =   495
         Width           =   1590
      End
      Begin VB.CheckBox chkFormats 
         Caption         =   "Excel"
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   1590
      End
   End
End
Attribute VB_Name = "frmExportFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bExport As Boolean

'Private Sub chkFormats_Click(Index As Integer)
'    Dim i As Integer
'    i = 0
'
'    If chkFormats(Index).Value = vbChecked Then
'
'        Do Until i = chkFormats.Count
'
'            If Not i = Index Then chkFormats(i).Value = vbUnchecked
'            i = i + 1
'        Loop
'
'    End If
'
'End Sub

Private Sub cmdCancel_Click()
    bExport = False
    Me.Hide
End Sub

Private Sub cmdExport_Click()
    bExport = True
    Me.Hide
End Sub

Private Sub Form_Load()
    If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If

End Sub

