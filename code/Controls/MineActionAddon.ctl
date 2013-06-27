VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.UserControl MineActionAddon 
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3195
   ScaleHeight     =   5280
   ScaleWidth      =   3195
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   5280
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3195
      _cx             =   5636
      _cy             =   9313
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   2
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"MineActionAddon.ctx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame FraMineAction 
         Caption         =   "Mine Action Themes"
         Height          =   3930
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   3015
         Begin VB.Frame FraSource 
            Caption         =   "IMSMA Source:"
            Height          =   1680
            Left            =   45
            TabIndex        =   8
            Top             =   315
            Width           =   2175
            Begin VB.CheckBox chkSource 
               Caption         =   "LIS"
               Height          =   150
               Index           =   4
               Left            =   135
               TabIndex        =   13
               Top             =   1395
               Width           =   1995
            End
            Begin VB.CheckBox chkSource 
               Caption         =   "MAG Erbil"
               Height          =   285
               Index           =   3
               Left            =   135
               TabIndex        =   12
               Top             =   1095
               Width           =   1995
            End
            Begin VB.CheckBox chkSource 
               Caption         =   "RMAC Central"
               Height          =   285
               Index           =   2
               Left            =   135
               TabIndex        =   11
               Top             =   810
               Width           =   1995
            End
            Begin VB.CheckBox chkSource 
               Caption         =   "RMAC South"
               Height          =   285
               Index           =   1
               Left            =   135
               TabIndex        =   10
               Top             =   510
               Width           =   1995
            End
            Begin VB.CheckBox chkSource 
               Caption         =   "IKMAA"
               Height          =   285
               Index           =   0
               Left            =   135
               TabIndex        =   9
               Top             =   225
               Width           =   1995
            End
         End
         Begin VB.CheckBox chkMineActionLyrs 
            Caption         =   "Hazards"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   7
            Top             =   2070
            Width           =   1680
         End
         Begin VB.CheckBox chkMineActionLyrs 
            Caption         =   "Dangerous Areas"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   6
            Top             =   2370
            Width           =   1680
         End
         Begin VB.CheckBox chkMineActionLyrs 
            Caption         =   "Mine Fields"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   5
            Top             =   2670
            Width           =   1680
         End
         Begin VB.CheckBox chkMineActionLyrs 
            Caption         =   "Mined Area"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   4
            Top             =   2985
            Width           =   1680
         End
         Begin VB.CheckBox chkMineActionLyrs 
            Caption         =   "Victims"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   3
            Top             =   3285
            Width           =   1680
         End
      End
      Begin VB.Frame FraTools 
         Caption         =   "Tools"
         Height          =   1110
         Left            =   90
         TabIndex        =   1
         Top             =   4080
         Width           =   3015
         Begin VB.CommandButton cmdUpdateFrom 
            Caption         =   "Update From Server"
            Height          =   465
            Left            =   90
            TabIndex        =   14
            Top             =   360
            Width           =   1230
         End
      End
   End
End
Attribute VB_Name = "MineActionAddon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_RSSettings As adodb.Recordset
Public Event LoadMALyr(oLyr As TatukGIS_XDK9.XGIS_LayerVector)
Public Event UnloadMALyr(sLyrName As String)

Public Sub Init(RSSettings As adodb.Recordset)
    Set m_RSSettings = RSSettings
End Sub

Private Sub chkMineActionLyrs_Click(Index As Integer)
    Dim sVal As String
    Dim oLyr As New TatukGIS_XDK9.XGIS_LayerVector

    Select Case Index
    
        Case 0
            sVal = "HazardLyr"

        Case 1
            sVal = "DaLyr"

        Case 2
            sVal = "MFLyr"

        Case 3
            sVal = "MALyr"

        Case 4
            sVal = "VictimsLyr"
    End Select
    
    SafeMoveFirst g_RSAppSettings
    g_RSAppSettings.Find "SettingName = '" & sVal & "'"
    
    If chkMineActionLyrs(Index).Value = vbChecked Then
        oLyr.Name = g_RSAppSettings.Fields.Item("SettingValue1").Value
        oLyr.Path = g_RSAppSettings.Fields.Item("SettingValue2").Value
        oLyr.Open
        RaiseEvent LoadMALyr(oLyr)
    Else
        RaiseEvent UnloadMALyr(g_RSAppSettings.Fields.Item("SettingValue1").Value)
    End If
End Sub
