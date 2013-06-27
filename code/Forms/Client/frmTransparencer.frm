VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransparencer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "OASIS Layer Transparency settings"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraTransparancy 
      Caption         =   "Transparency:"
      Height          =   645
      Left            =   30
      TabIndex        =   2
      Top             =   720
      Width           =   3915
      Begin MSComctlLib.Slider scrTransparency 
         Height          =   255
         Left            =   750
         TabIndex        =   3
         Top             =   270
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   100
         TickStyle       =   3
         Value           =   100
      End
      Begin VB.Label lblMax 
         Caption         =   "100%"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblmin 
         Caption         =   "0%"
         Height          =   255
         Left            =   3330
         TabIndex        =   4
         Top             =   240
         Width           =   285
      End
   End
   Begin VB.Frame FraChooseLayer 
      Caption         =   "Choose Layer:"
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3945
      Begin VB.ComboBox ComLayers 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmTransparencer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oGIS As TatukGIS_XDK9.XGIS_Viewer
Private LoadDone As Boolean
Private LyrCol As Collection

Public Sub Init(GISVwr As TatukGIS_XDK9.XGIS_Viewer)
Dim i As Integer

    Set LyrCol = New Collection
    
    Set oGIS = GISVwr
    
    For i = 0 To oGIS.Items.Count - 1
        LyrCol.Add oGIS.Items.Item(i).Name, oGIS.Items.Item(i).caption
        ComLayers.AddItem oGIS.Items.Item(i).caption
    Next
    
    If ComLayers.ListCount > 0 Then ComLayers.ListIndex = 0
    
    LoadDone = True

End Sub

Private Sub scrTransparency_Change()
   Dim lL As TatukGIS_XDK9.XGIS_LayerVector
   
   If Not LoadDone Then Exit Sub
   
   Set lL = oGIS.get(LyrCol.Item(ComLayers.List(ComLayers.ListIndex)))
   
   If lL Is Nothing Then Exit Sub
   ' change transparency
   lL.Transparency = scrTransparency.Value
   oGIS.UpDate
End Sub
