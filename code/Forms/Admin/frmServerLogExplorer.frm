VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmServerLogExplorer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Server Log Explorer"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13800
   Icon            =   "frmServerLogExplorer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   13800
   StartUpPosition =   1  'CenterOwner
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13800
      _cx             =   24342
      _cy             =   14843
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
      BorderWidth     =   3
      ChildSpacing    =   2
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
      Begin VB.CommandButton cmdCheckServerLog 
         Caption         =   "Check Server Log"
         Height          =   285
         Left            =   10350
         TabIndex        =   2
         Top             =   60
         Width           =   1635
      End
      Begin VB.TextBox txtServerURLLog 
         Height          =   285
         Left            =   540
         TabIndex        =   1
         Text            =   "http://www.immap.org/"
         Top             =   60
         Width           =   9735
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   7935
         Left            =   0
         TabIndex        =   3
         Top             =   405
         Width           =   13650
         ExtentX         =   24077
         ExtentY         =   13996
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Label lblURL1 
         Caption         =   "URL:"
         Height          =   240
         Left            =   0
         TabIndex        =   4
         Top             =   60
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmServerLogExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheckServerLog_Click()
        '<EhHeader>
        On Error GoTo cmdCheckServerLog_Click_Err
        '</EhHeader>
        
        If Right$(txtServerURLLog.Text, 1) <> "/" Then
            txtServerURLLog.Text = txtServerURLLog.Text & "/"
        End If

100     WebBrowser1.Navigate txtServerURLLog.Text & "oasis.asp?logger=1"
    
        '<EhFooter>
        Exit Sub

cmdCheckServerLog_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmServerLogExplorer.cmdCheckServerLog_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     txtServerURLLog.Text = frmDatabaseConnect.txtServerURL
'GetSetting(App.EXEName, "Settings", "WebServerURLLog", "http://www.immap.org/")
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmServerLogExplorer.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
