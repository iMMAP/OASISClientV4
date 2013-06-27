VERSION 5.00
Begin VB.Form frmOGISFormats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OGIS DB Formats"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   2415
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptOGisFormat 
      Caption         =   "Open Gis Blob 2"
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   6
      ToolTipText     =   "OASIS will look for SPATIAL_REFERENCE_SYS table"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   60
      TabIndex        =   5
      Top             =   1980
      Width           =   2295
   End
   Begin VB.Frame FraChooseOGIS 
      Caption         =   "Choose OGIS Format:"
      Height          =   1935
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.OptionButton OptOGisFormat 
         Caption         =   "Open Gis Normalized 2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "OASIS will look for SPATIAL_REFERENCE_SYS table"
         Top             =   1260
         Width           =   1935
      End
      Begin VB.OptionButton OptOGisFormat 
         Caption         =   "Open Gis Normalized"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "OASIS will look for SPATIAL_REFERENCE_SYSTEMS"
         Top             =   940
         Width           =   1935
      End
      Begin VB.OptionButton OptOGisFormat 
         Caption         =   "Open Gis Wkt"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Normal OGIS WKT table"
         Top             =   620
         Width           =   1935
      End
      Begin VB.OptionButton OptOGisFormat 
         Caption         =   "Open Gis Blob"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "OASIS will look for SPATIAL_REFERENCE_SYS table"
         Top             =   300
         Value           =   -1  'True
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmOGISFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    g_sOGISFormat = "OpenGisBlob"
End Sub

Private Sub OptOGisFormat_Click(Index As Integer)
    Select Case Index
    
        Case 0
            g_sOGISFormat = "OpenGisBlob"
        Case 1
            g_sOGISFormat = "OpenGisWkt"
        Case 2
            g_sOGISFormat = "OpenGisNormalized"
        Case 3
            g_sOGISFormat = "OpenGisNormalized2"
        Case Else
            g_sOGISFormat = "OpenGisBlob2"
    End Select
End Sub
