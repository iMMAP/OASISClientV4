VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVillageAvailable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Available Places"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2955
   Icon            =   "frmVillageAvailable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   2955
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvPlaces 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   7646
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
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmVillageAvailable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AddPlaces(colPlaces As Collection)
        '<EhHeader>
        On Error GoTo AddPlaces_Err
        '</EhHeader>
    Dim i As Integer

100     lvPlaces.ListItems.Clear
        'lvPlaces.ColumnHeaders.Add Text:="Name"
     
    
102     For i = 1 To colPlaces.Count
104         lvPlaces.ListItems.Add Text:=colPlaces.Item(i)
            'lvPlaces.ListItems.Item(0).ListSubItems.Add colPlaces.Item(i)
        Next
        '<EhFooter>
        Exit Sub

AddPlaces_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmVillageAvailable.AddPlaces " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub Form_Load()
If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
End Sub
