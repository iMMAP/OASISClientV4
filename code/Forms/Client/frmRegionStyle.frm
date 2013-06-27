VERSION 5.00
Begin VB.Form frmRegionStyle 
   Caption         =   "Region Style"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3480
   LinkTopic       =   "Form2"
   ScaleHeight     =   5835
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   240
      Left            =   1665
      TabIndex        =   4
      Top             =   5400
      Width           =   825
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   240
      Left            =   585
      TabIndex        =   3
      Top             =   5400
      Width           =   780
   End
   Begin VB.Frame FraSample 
      Caption         =   "Sample"
      Height          =   735
      Left            =   270
      TabIndex        =   2
      Top             =   4500
      Width           =   2805
   End
   Begin VB.Frame FraBorder 
      Caption         =   "Border"
      Height          =   2715
      Left            =   225
      TabIndex        =   1
      Top             =   1575
      Width           =   2895
   End
   Begin VB.Frame FraFill 
      Caption         =   "Fill"
      Height          =   1275
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   2895
      Begin VB.Label lblPattern 
         Caption         =   "Pattern:"
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   270
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmRegionStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oShp As TatukGIS_XDK9.XGIS_Shape

Public Sub Init(oshp As TatukGIS_XDK9.XGIS_Shape)
    Set m_oShp = oshp
End Sub

Private Sub SetStyle()
'm_oSHP.MakeEditable = True
    With m_oShp.Params.area
        .Color = RGB(0, 0, 255)
        '.OutlineColor = RGB(255, 0, 255)
        '.OutlinePattern = XbsSolid
        '.OutlineStyle = XpsDashDot
        '.OutlineWidth = 3
      '  .Pattern = XbsHorizontal
    End With
    
    
End Sub

Private Sub cmdOk_Click()
    SetStyle
    Me.Hide
End Sub

Private Sub Form_Load()
If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
End Sub
