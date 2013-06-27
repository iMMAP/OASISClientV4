VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWordBookMarks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Available Export Locations"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3210
   Icon            =   "frmWordBookMarks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraChooseExport 
      Caption         =   "Choose Export Areas:"
      Height          =   4425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3210
      Begin MSComctlLib.ListView lvBookMarks 
         Height          =   3930
         Left            =   135
         TabIndex        =   1
         Top             =   360
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   6932
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Export Section"
            Object.Width           =   4304
         EndProperty
      End
   End
End
Attribute VB_Name = "frmWordBookMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ClosingDown(iBkrMark As Integer, bChoosenMark)

Private Sub Form_Load()
If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
    For i = 1 To lvBookMarks.ListItems.Count
        If lvBookMarks.ListItems.Item(i).Checked Then
            RaiseEvent ClosingDown(i, True)
            Exit Sub
        End If
    Next
    
    
    RaiseEvent ClosingDown(0, False)
End Sub
