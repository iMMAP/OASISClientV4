VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmSQLchecker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS SQL Tester"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9930
   Icon            =   "frmSQLchecker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9930
   StartUpPosition =   1  'CenterOwner
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6840
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9930
      _cx             =   17515
      _cy             =   12065
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
      Begin VB.TextBox txtURL 
         Height          =   330
         Left            =   765
         TabIndex        =   4
         Text            =   "http://www.immap.org/"
         Top             =   2250
         Width           =   7755
      End
      Begin VB.CommandButton cmdTestSQL 
         Caption         =   "Run SQL"
         Height          =   375
         Left            =   8640
         TabIndex        =   3
         Top             =   2205
         Width           =   1185
      End
      Begin VB.TextBox txtResult 
         ForeColor       =   &H00FF0000&
         Height          =   3750
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   3015
         Width           =   9645
      End
      Begin VB.TextBox txtSQL 
         ForeColor       =   &H0000C000&
         Height          =   1725
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmSQLchecker.frx":6852
         Top             =   405
         Width           =   9690
      End
      Begin VB.Label lblSQL 
         Caption         =   "SQL:"
         Height          =   330
         Left            =   135
         TabIndex        =   7
         Top             =   135
         Width           =   780
      End
      Begin VB.Label lblServerResult 
         Caption         =   "Server Result:"
         Height          =   285
         Left            =   135
         TabIndex        =   6
         Top             =   2790
         Width           =   1500
      End
      Begin VB.Label lblServerURL 
         Caption         =   "Server URL:"
         Height          =   420
         Left            =   135
         TabIndex        =   5
         Top             =   2160
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmSQLchecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTestSQL_Click()
        '<EhHeader>
        On Error GoTo cmdTestSQL_Click_Err
        '</EhHeader>
        Dim sSQL As String
        Dim sReturnValue As String
        
100     txtResult.Text = ""

102     sSQL = txtURL.Text & "Oasis.asp?testsql=" & CheckEncrypt(txtSQL.Text)
104     sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(sSQL, True)

106     txtResult.Text = sReturnValue

        '<EhFooter>
        Exit Sub

cmdTestSQL_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSQLchecker.cmdTestSQL_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100    txtURL.Text = WebSite 'GetSetting(App.EXEName, "Settings", "WebServerSQLCheck", "http://www.immap.org/")

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSQLchecker.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     SaveSetting App.EXEName, "Settings", "WebServerSQLCheck", txtURL.Text

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmSQLchecker.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

