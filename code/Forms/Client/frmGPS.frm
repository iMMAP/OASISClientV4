VERSION 5.00
Object = "{9C989D2F-3596-477B-B719-6DC4E6893AF2}#1.0#0"; "TatukGIS_XDK10_WIN32.ocx"
Begin VB.Form frmGPS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS GPS Utility"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4260
   Icon            =   "frmGPS.frx":0000
   LinkTopic       =   "OASIS GPS Tracker"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraGPS 
      Caption         =   "GPS"
      Height          =   3615
      Left            =   -90
      TabIndex        =   0
      Top             =   -180
      Width           =   4515
      Begin TatukGIS_XDK10.XGIS_GpsNmea GPS 
         Height          =   2895
         Left            =   180
         TabIndex        =   5
         Top             =   600
         Width           =   4095
         Timeout         =   1000
         Com             =   1
         BaudRate        =   4800
         Align           =   2
         BevelInner      =   0
         BevelOuter      =   0
         BorderStyle     =   0
         Color           =   -16777201
         Ctl3D           =   0   'False
         Enabled         =   -1  'True
         FullRepaint     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   0   'False
         Object.Visible         =   -1  'True
         DoubleBuffered  =   0   'False
      End
      Begin VB.ComboBox ComBAUD 
         Height          =   315
         ItemData        =   "frmGPS.frx":6852
         Left            =   180
         List            =   "frmGPS.frx":6883
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   1785
      End
      Begin VB.CommandButton cmdActivate 
         Caption         =   "Activate"
         Height          =   315
         Left            =   3285
         TabIndex        =   2
         Top             =   225
         Width           =   885
      End
      Begin VB.TextBox txtComPort 
         Height          =   285
         Left            =   2475
         TabIndex        =   1
         Text            =   "1"
         Top             =   270
         Width           =   765
      End
      Begin VB.Label lblCOMPort 
         Caption         =   "COM Port:"
         Height          =   435
         Left            =   2025
         TabIndex        =   4
         Top             =   225
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmGPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCommand6_Click()
    'OASISLocator1.Init GIS.viewer
End Sub

Private Sub ComBAUD_Click()
    GPS.BaudRate = ComBAUD.List(ComBAUD.ListIndex)
End Sub

Private Sub cmdActivate_Click()
        '<EhHeader>
        On Error GoTo cmdActivate_Click_Err
        '</EhHeader>
100     If cmdActivate.caption = "Activate" Then
102         cmdActivate.caption = "Stop GPS"
104         GPS.Com = CInt(txtComPort.Text)
106         If IsNumeric(GPS.BaudRate = ComBAUD.List(ComBAUD.ListIndex)) Then
108             GPS.BaudRate = CInt(ComBAUD.List(ComBAUD.ListIndex))
            Else
110             GPS.BaudRate = 1
            End If
        
112         GPS.Active = True
            GPS.InitiateAction
        Else
114         cmdActivate.caption = "Activate"
116         GPS.Active = False
        End If
        '<EhFooter>
        Exit Sub

cmdActivate_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmGPS.cmdActivate_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub GPS_OnNmea(translated As Boolean, Name As String, ByVal Items As Object, ByVal parsed As Boolean)
    Me.caption = Name & GPS.Longitude & "  " & GPS.Latitude
End Sub

Private Sub GPS_OnPosition(translated As Boolean)
    Me.caption = GPS.Longitude & " Lat: " & GPS.Latitude
    'gps.InitiateAction
End Sub
