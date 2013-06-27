VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmCommodity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commodities"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5805
   Icon            =   "frmCommodity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCommodity.frx":6852
   ScaleHeight     =   7515
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elmain 
      Height          =   7515
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5805
      _cx             =   10239
      _cy             =   13256
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
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   2
      ChildSpacing    =   1
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
      _GridInfo       =   $"frmCommodity.frx":D0A4
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridCommodity 
         Height          =   6765
         Left            =   30
         OleObjectBlob   =   "frmCommodity.frx":D0E7
         TabIndex        =   1
         Top             =   30
         Width           =   5745
      End
   End
End
Attribute VB_Name = "frmCommodity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Init()
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
100     With dxDBGridCommodity
102         .Event = 1 'EGOnCustomDrawCell
104         .EventEnabled = True
106         .Options.Set 18 'egoAutoWidth
108         .Columns.ApplyBestFit Nothing
110         .Dataset.ADODataset.ConnectionString = m_Cnn.ConnectionString
        
112         SafeMoveFirst g_RSAppSettings
114         g_RSAppSettings.Find "SettingName = 'CommodityAddonDataSQL'"

116         .Dataset.ADODataset.CommandText = g_RSAppSettings.Fields.Item("SettingValue1").Value '"SELECT ID, Impact, Town, Province, District FROM Scoring ORDER BY Scoring DESC"
118         .Dataset.ADODataset.CommandType = cmdText
120         .Dataset.Open

122         .KeyField = g_RSAppSettings.Fields.Item("SettingValue2").Value
124         .Dataset.Active = True
126         .Columns.RetrieveFields
128         .Dataset.ADODataset.Requery
130         .Columns.Item(0).Visible = False 'ID
132         .m.AddGroupColumn .Columns.Item(1)
134         .m.AddGroupColumn .Columns.Item(3)
136         .m.AddGroupColumn .Columns.Item(4)
        
        End With

        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmCommodity.init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
    If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If

End Sub
