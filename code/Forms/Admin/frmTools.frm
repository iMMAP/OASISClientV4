VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Begin VB.Form frmAdminTools 
   Caption         =   "Administration Tools"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7965
   Icon            =   "frmTools.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4635
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Elastic c1Utils 
      Height          =   4635
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7965
      _cx             =   14049
      _cy             =   8176
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
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
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
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin CONTROLSLibCtl.dxPicBtn cmdServerDateTime 
         Height          =   720
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmTools.frx":6852
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn dxEncryptTool 
         Height          =   720
         Left            =   2520
         TabIndex        =   3
         Top             =   600
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmTools.frx":75A4
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdServerSettings 
         Height          =   720
         Left            =   720
         TabIndex        =   5
         Top             =   1800
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmTools.frx":82F6
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdConnString 
         Height          =   720
         Left            =   2520
         TabIndex        =   7
         Top             =   2040
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmTools.frx":9048
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdQueryTester 
         Height          =   720
         Left            =   4440
         TabIndex        =   9
         Top             =   1800
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmTools.frx":9D9A
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdLogExplorer 
         Height          =   720
         Left            =   6480
         TabIndex        =   11
         Top             =   600
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmTools.frx":AAEC
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdSynchExplorer 
         Height          =   720
         Left            =   6480
         TabIndex        =   12
         Top             =   2040
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmTools.frx":B83E
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdOASISASPrev 
         Height          =   720
         Left            =   720
         TabIndex        =   15
         Top             =   3240
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmTools.frx":C590
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn dxIncFont 
         Height          =   720
         Left            =   2490
         TabIndex        =   17
         Top             =   3420
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmTools.frx":D2E2
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn dxPicBtnPixelStore 
         Height          =   480
         Left            =   4560
         TabIndex        =   19
         Top             =   3330
         Width           =   480
         _Version        =   65536
         _cx             =   847
         _cy             =   847
         Picture         =   "frmTools.frx":E034
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin CONTROLSLibCtl.dxPicBtn cmdIncidentsMaint 
         Height          =   720
         Left            =   4440
         TabIndex        =   21
         Top             =   330
         Width           =   720
         _Version        =   65536
         _cx             =   1270
         _cy             =   1270
         Picture         =   "frmTools.frx":E886
         BackColor       =   15790320
         Enabled         =   -1  'True
         Style           =   3
         DitherStyle     =   0
         DitherColor     =   255
         GroupIndex      =   -1
         Stuck           =   0   'False
         Pushed          =   0   'False
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Incidents Maintainance"
         Height          =   375
         Left            =   4080
         TabIndex        =   22
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Export To Pixel Database"
         Height          =   435
         Left            =   4170
         TabIndex        =   20
         Top             =   3990
         Width           =   1305
      End
      Begin VB.Label lblOASISIncident 
         Alignment       =   2  'Center
         Caption         =   "Incident Symbols"
         Height          =   195
         Left            =   2250
         TabIndex        =   18
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "OASIS.ASP server revision"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Log Explorer"
         Height          =   375
         Left            =   6120
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Synch Explorer"
         Height          =   375
         Left            =   6120
         TabIndex        =   13
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Query Tester"
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Connection String"
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Server Settings"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Encryption Tool"
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Server Time / Date"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAdminTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdConnString_Click()

    If frmConnectionString.Visible Then Exit Sub
    frmConnectionString.Show vbModeless, Me

End Sub

'Private Sub cmdDataPackExp_Click()
'
'    If frmOASISSynch.Visible Then Exit Sub
'    frmOASISSynch.Show vbModeless, Me
'End Sub

Private Sub cmdIncidentsMaint_Click()
    If frmIncidentsMaintainance.Visible Then Exit Sub
    frmIncidentsMaintainance.Show vbModeless, Me
End Sub

Private Sub cmdLogExplorer_Click()

    'If Not frmLOG.Visible Then frmLOG.Show vbModeless, Me
    If frmServerLogExplorer.Visible Then Exit Sub
    frmServerLogExplorer.Show vbModeless, Me
End Sub

Private Sub cmdOASISASPrev_Click()
    
    MsgBox getASPFileVersionAndDate(WebSite)
    
End Sub

Private Sub cmdServerDateTime_Click()

    If frmTimeDateTest.Visible Then Exit Sub
    frmTimeDateTest.Show vbModeless, Me
End Sub

Private Sub cmdServerSettings_Click()

    If frmServerStatus.Visible Then Exit Sub
    frmServerStatus.Show vbModeless, Me
End Sub

Private Sub cmdSynchExplorer_Click()

    '    If oSynchExplorer Is Nothing Then
    '        Set oSynchExplorer = CreateObject("Synch_Explorer.clsExplorer")
    '    End If
    '
    '    oSynchExplorer.ShowExplorer Me.hwnd
    
    MsgBox "This form is under construction!", vbInformation, "Not available"
    Exit Sub
    If frmSynchReader.Visible Then Exit Sub
    frmSynchReader.Show vbModeless, Me
End Sub

Private Sub dxEncryptTool_Click()

    If frmCiphers.Visible Then Exit Sub
    frmCiphers.Show vbModeless, Me
End Sub

Private Sub cmdQueryTester_Click()

    If frmSQLchecker.Visible Then Exit Sub
    frmSQLchecker.Show vbModeless, Me
End Sub

Private Sub dxIncFont_Click()
    On Error Resume Next
    Shell App.Path & "\OASIS_Fonts.exe", vbNormalFocus
End Sub

Private Sub dxPicBtnPixelStore_Click()
    MsgBox "Functionality not implemented yet!"
End Sub
