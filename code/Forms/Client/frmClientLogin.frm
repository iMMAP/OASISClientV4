VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1950
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4335
   ClipControls    =   0   'False
   Icon            =   "frmClientLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Tab C1Tab 
      Height          =   2035
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   5295
      _cx             =   9340
      _cy             =   3590
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "Login|Database|Server|Reset|Password"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   5
      Position        =   6
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   1950
         Left            =   6840
         TabIndex        =   28
         Top             =   45
         Width           =   4275
         Begin VB.CommandButton cmdUpdatePassword 
            Caption         =   "Update Password"
            Enabled         =   0   'False
            Height          =   315
            Left            =   0
            TabIndex        =   29
            Top             =   1440
            Width           =   4140
         End
         Begin XpressEditorsLibCtl.dxMaskEdit dxMaskEditOldPass 
            Height          =   315
            Left            =   1380
            OleObjectBlob   =   "frmClientLogin.frx":6852
            TabIndex        =   30
            Top             =   375
            Width           =   2760
         End
         Begin XpressEditorsLibCtl.dxMaskEdit dxMaskEditNewPass1 
            Height          =   315
            Left            =   1380
            OleObjectBlob   =   "frmClientLogin.frx":68CD
            TabIndex        =   31
            Top             =   690
            Width           =   2760
         End
         Begin XpressEditorsLibCtl.dxMaskEdit dxMaskEditNewPass2 
            Height          =   315
            Left            =   1380
            OleObjectBlob   =   "frmClientLogin.frx":6948
            TabIndex        =   32
            Top             =   1005
            Width           =   2760
         End
         Begin C1SizerLibCtl.C1Elastic C1EChangeUser 
            Height          =   315
            Left            =   0
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   0
            Width           =   4140
            _cx             =   7303
            _cy             =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   16711680
            ForeColor       =   -2147483634
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "  Change User Password"
            Align           =   0
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
            FrameWidth      =   0
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
         Begin C1SizerLibCtl.C1Elastic C1EOld 
            Height          =   315
            Left            =   0
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   375
            Width           =   1380
            _cx             =   2434
            _cy             =   556
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
            Caption         =   "Old:"
            Align           =   0
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
            FrameWidth      =   0
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
         Begin C1SizerLibCtl.C1Elastic C1ENew1 
            Height          =   315
            Left            =   0
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   690
            Width           =   1380
            _cx             =   2434
            _cy             =   556
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
            Caption         =   "New (1):"
            Align           =   0
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
            FrameWidth      =   0
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
         Begin C1SizerLibCtl.C1Elastic C1ENew2 
            Height          =   315
            Left            =   0
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1005
            Width           =   1380
            _cx             =   2434
            _cy             =   556
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
            Caption         =   "New (2):"
            Align           =   0
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
            FrameWidth      =   0
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   1950
         Left            =   6540
         TabIndex        =   26
         Top             =   45
         Width           =   4275
         Begin VB.CommandButton cmdRestoreSQL 
            Caption         =   "Restore SQL Server Backup"
            Height          =   375
            Left            =   180
            TabIndex        =   38
            Top             =   960
            Width           =   3975
         End
         Begin VB.CommandButton cmdReloadProfile 
            Caption         =   "Reload Profile Settings"
            Height          =   420
            Left            =   180
            TabIndex        =   27
            Top             =   420
            Width           =   3960
         End
         Begin C1SizerLibCtl.C1Elastic C1EOASISReset 
            Height          =   315
            Left            =   0
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   0
            Width           =   4140
            _cx             =   7303
            _cy             =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   16711680
            ForeColor       =   -2147483634
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "  OASIS Reset Options"
            Align           =   0
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
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1950
         Left            =   5940
         TabIndex        =   20
         Top             =   45
         Width           =   4275
         Begin VB.Frame FraDatabase 
            BorderStyle     =   0  'None
            Height          =   1260
            Left            =   60
            TabIndex        =   21
            Top             =   360
            Width           =   4020
            Begin VB.OptionButton OptMicrosoftAccess 
               Caption         =   "Microsoft Access"
               Height          =   255
               Left            =   180
               TabIndex        =   24
               Top             =   120
               Value           =   -1  'True
               Width           =   1995
            End
            Begin VB.OptionButton OptMicrosoftSQL 
               Caption         =   "Microsoft SQL Server"
               Enabled         =   0   'False
               Height          =   195
               Left            =   180
               TabIndex        =   23
               Top             =   540
               Width           =   2775
            End
            Begin VB.OptionButton OptPostgreSQL 
               Caption         =   "PostgreSQL"
               Enabled         =   0   'False
               Height          =   195
               Left            =   180
               TabIndex        =   22
               Top             =   900
               Width           =   2775
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1ESelectDatabase 
            Height          =   315
            Left            =   0
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   4140
            _cx             =   7303
            _cy             =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   16711680
            ForeColor       =   -2147483634
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "  Select Database"
            Align           =   0
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
            FrameWidth      =   0
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1950
         Left            =   6240
         TabIndex        =   15
         Top             =   45
         Width           =   4275
         Begin VB.CommandButton cmdConnect 
            Caption         =   "Connect"
            Enabled         =   0   'False
            Height          =   315
            Left            =   0
            TabIndex        =   19
            Top             =   1500
            Width           =   2070
         End
         Begin VB.CommandButton cmdAddManually 
            Caption         =   "Add Manually"
            Height          =   315
            Left            =   2070
            TabIndex        =   18
            Top             =   1500
            Width           =   2070
         End
         Begin VB.ListBox listServer 
            Height          =   1035
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   300
            Width           =   4140
         End
         Begin C1SizerLibCtl.C1Elastic C1EOASISCloud 
            Height          =   315
            Left            =   0
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   4140
            _cx             =   7303
            _cy             =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   16711680
            ForeColor       =   -2147483634
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "  OASIS Cloud Public Servers"
            Align           =   0
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
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1950
         Left            =   45
         TabIndex        =   7
         Top             =   45
         Width           =   4275
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            Default         =   -1  'True
            Height          =   315
            Left            =   0
            TabIndex        =   3
            Top             =   1560
            Width           =   1860
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "Cancel"
            Height          =   315
            Left            =   2460
            TabIndex        =   4
            Top             =   1560
            Width           =   1800
         End
         Begin VB.CheckBox chkRememberUser 
            Caption         =   "Remember me"
            Height          =   315
            Left            =   1860
            TabIndex        =   5
            Top             =   1260
            Width           =   1800
         End
         Begin VB.CheckBox chkWorkOnline 
            Caption         =   "Work online"
            Height          =   315
            Left            =   0
            TabIndex        =   2
            Top             =   1260
            Value           =   1  'Checked
            Width           =   1860
         End
         Begin VB.CommandButton cmdExpand 
            Height          =   315
            Left            =   3660
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmClientLogin.frx":69C3
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1260
            Width           =   600
         End
         Begin VB.TextBox txtUserName 
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            Top             =   0
            Width           =   3000
         End
         Begin VB.TextBox txtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1260
            PasswordChar    =   "*"
            TabIndex        =   1
            Top             =   315
            Width           =   3000
         End
         Begin VB.ComboBox ComServer 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmClientLogin.frx":6D61
            Left            =   1260
            List            =   "frmClientLogin.frx":6D68
            Style           =   1  'Simple Combo
            TabIndex        =   9
            Text            =   "ComServer"
            Top             =   630
            Width           =   3000
         End
         Begin VB.ComboBox ComDatabase 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmClientLogin.frx":6D75
            Left            =   1260
            List            =   "frmClientLogin.frx":6D7C
            Style           =   1  'Simple Combo
            TabIndex        =   8
            Text            =   "Microsoft Access"
            Top             =   945
            Width           =   3000
         End
         Begin VB.Label lblUserName 
            Caption         =   "&User Name:"
            Height          =   315
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1260
         End
         Begin VB.Label lblPassword 
            Caption         =   "&Password:"
            Height          =   315
            Left            =   0
            TabIndex        =   12
            Top             =   315
            Width           =   1260
         End
         Begin VB.Label lblServer 
            Caption         =   "Server:"
            Height          =   315
            Left            =   0
            TabIndex        =   11
            Top             =   630
            Width           =   1260
         End
         Begin VB.Label lblDatabase 
            Caption         =   "Database:"
            Height          =   315
            Left            =   0
            TabIndex        =   10
            Top             =   945
            Width           =   1260
         End
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" (ByVal lpBuffer As String, _
                                      nSize As Long) As Long
Private Declare Sub SleepAPI Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Private m_bPrevLogin As Boolean
Private m_LoginParams() As String
Private m_Login As New clsTooltips
Private m_bChangeDefaultUser As Boolean
Private m_RSProfile As New ADODB.Recordset
Private m_bInternetOK As Boolean
Private WithEvents m_HK As cRegHotKey
Attribute m_HK.VB_VarHelpID = -1
Public OK As Boolean
Private m_bInitReady As Boolean

'Private MsXmlHttp As New MSXML2.XML
Private WithEvents m_frmUpdateSettings As frmUpdateSettings
Attribute m_frmUpdateSettings.VB_VarHelpID = -1
Private WithEvents m_frmForcedLogin As frmForcedLogin
Attribute m_frmForcedLogin.VB_VarHelpID = -1
Public Event Done()
Private sInitialMSACCESSServer As String
Private lStartWidth As Long
Private bOASISv3Login As Boolean

 
Private Sub c1Tab_Click()
If c1Tab.CurrTab = 4 Then
    
        C1EChangeUser.caption = "  Change password for user: " & txtUserName.Text
    
    End If
End Sub

Private Sub cmdAddManually_Click()

    Dim Wipe As VbMsgBoxResult
    Dim CN As ADODB.Connection
    Me.Hide

    Dim sString As String
    sString = InputBox("Please enter in the server address (without 'http://')", "Custom Server Address")

    If sString = ComServer.Text Then
        MsgBox "This is the existing server!"
    Else

        If Len(sString) > 0 Then
        
            If sString <> sInitialMSACCESSServer And OptMicrosoftAccess.value = True Then

                Wipe = MsgBox("You are changing server.  In order to accomplish this the client database must be wiped.  Are you sure you want to proceed?", vbYesNo, "Wipe database?")

                If Wipe = vbYes Then
        
                    Set CN = New ADODB.Connection
                    CN.Open GetConnectionString(g_sAppPath & "\data\db\OasisClient.mdb")
                    WipeTablesForReset CN
                    CN.Close
                    Set CN = Nothing
      
                    sInitialMSACCESSServer = sString
        
                    ComServer.Text = sString

                    If Len(listServer.Text) > 0 Then cmdConnect.Enabled = True
                    MsgBox "The OASIS Server has been updated", vbInformation
                End If

            Else
        
                ComServer.Text = sString

                If Len(listServer.Text) > 0 Then cmdConnect.Enabled = True
                MsgBox "The OASIS Server has been updated", vbInformation
    
            End If
        
        End If
        
    End If

    Me.Show
    g_sAppServerPath = "http://" & ComServer.Text
    cmdRestoreSQL.Enabled = IIf(FileExists(g_sAppPath & "\data\db\" & Replace(g_sAppServerPath, "/", "-") & ".bak"), True, False)
        
        If ComServer.Text = "atlantis.oasiswebservice.org" Then
        txtUserName = "demo"
        txtPassword = "demo"
      '  txtPassword.Enabled = False
      '  txtUserName.Enabled = False
        
    Else
        txtPassword.Enabled = True
        txtUserName.Enabled = True
    End If

End Sub

Private Sub cmdConnect_Click()
        '<EhHeader>
        On Error GoTo cmdConnect_Click_Err
        '</EhHeader>

        Dim Wipe As VbMsgBoxResult
        Dim CN As ADODB.Connection
100     Me.Hide

102     If listServer.Text <> sInitialMSACCESSServer And OptMicrosoftAccess.value = True Then

104         Wipe = MsgBox("You are changing server.  In order to accomplish this the client database must be wiped.  Are you sure you want to proceed?", vbYesNo, "Wipe database?")

106         If Wipe = vbYes Then
        
108             Set CN = New ADODB.Connection
110             CN.Open GetConnectionString(g_sAppPath & "\data\db\OasisClient.mdb")
112             WipeTablesForReset CN
114             CN.Close
116             Set CN = Nothing
            
118             ComServer.Text = listServer.Text
120             sInitialMSACCESSServer = ComServer.Text
122             MsgBox "The OASIS Server has been updated", vbInformation
124             c1Tab.CurrTab = 0
            End If

        Else
126         ComServer.Text = listServer.Text
128         MsgBox "The OASIS Server has been updated", vbInformation
130         c1Tab.CurrTab = 0
        End If
    
132     If ComServer.Text = "atlantis.oasiswebservice.org" Then
134         txtUserName = "demo"
136         txtPassword = "demo"
            'txtPassword.Enabled = False
           ' txtUserName.Enabled = False
        
        Else
138         txtPassword.Enabled = True
140         txtUserName.Enabled = True
        End If

142     Me.Show
144     Set CN = Nothing
    
146     g_sAppServerPath = "http://" & ComServer.Text
148     cmdRestoreSQL.Enabled = IIf(FileExists(g_sAppPath & "\data\db\" & Replace(g_sAppServerPath, "/", "-") & ".bak"), True, False)

        '<EhFooter>
        Exit Sub

cmdConnect_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLogin.cmdConnect_Click " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub cmdExpand_Click()

    Dim i As Long
    Dim sServers() As String
    
    If Me.Width = lStartWidth Then
    
        Me.Width = c1Tab.Width + 100

        If m_bInternetOK And listServer.ListCount = 0 Then
   
            listServer.Enabled = True
            'sServers = Split(OpenSilentHttpCommsResponse("http://www.oasiswebservice.org/getservers.txt", True), ";")
            sServers = Split(OpenServerResponseCompressed("http://" & g_sAppServerPath & "/oasis4.asp", "getservers", ""), ";")
            listServer.Clear

            Do Until i > UBound(sServers)
                listServer.AddItem sServers(i)
                i = i + 1
            Loop

            listServer.Text = ComServer.Text
    
        End If

    Else
        Me.Width = lStartWidth
    
    End If

End Sub

Private Function FileExists(Filename As String) As Integer
    Dim i As Integer
    
    On Local Error Resume Next
    i = Len(Dir$(Filename$))
    If Err Or i = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If
    On Local Error GoTo 0
End Function

Private Function FolderExists(sFolderName As String) As Integer
    
    Dim fso As New FileSystemObject
    If fso.FolderExists(sFolderName) Then
    FolderExists = True
    Else
    FolderExists = False
    End If
    
End Function

Private Function VerifyDataFolderIntegrity() As Boolean
        '<EhHeader>
        On Error GoTo VerifyDataFolderIntegrity_Err
        '</EhHeader>

        Dim sOutput As String
100     VerifyDataFolderIntegrity = True
    
        Dim i As Integer
        Dim sPaths(17) As String

102     sPaths(0) = g_sAppPath & "\data"
104     sPaths(1) = g_sAppPath & "\data\db"
106     sPaths(2) = g_sAppPath & "\data\gis"
108     sPaths(3) = g_sAppPath & "\data\sync"
110     sPaths(4) = g_sAppPath & "\data\templates"
112     sPaths(5) = g_sAppPath & "\data\user"
114     sPaths(6) = g_sAppPath & "\data\db\dynamicdata"
116     sPaths(7) = g_sAppPath & "\data\sync\import"
118     sPaths(8) = g_sAppPath & "\data\templates\ChartTemplates"
120     sPaths(9) = g_sAppPath & "\data\templates\FixedTemplates"
122     sPaths(10) = g_sAppPath & "\data\templates\printtemplates"
124     sPaths(11) = g_sAppPath & "\data\templates\spatialanalysis"
126     sPaths(12) = g_sAppPath & "\data\templates\SecurityChartTemplates"
128     sPaths(13) = g_sAppPath & "\data\user\Exports"
130     sPaths(14) = g_sAppPath & "\data\user\Maps"
132     sPaths(15) = g_sAppPath & "\data\user\Sessions"
134     sPaths(16) = g_sAppPath & "\data\user\utils"

136     i = 0

138     Do Until i = 17
    
140         If FolderExists(sPaths(i)) Then
                'sOutput = sOutput & "Folder [" & sPaths(i) & "] exists" & Chr(13)
            Else
                'sOutput = sOutput & "Folder [" & sPaths(i) & "] does not exist!" & Chr(13) & Chr(13)
142             MkDir sPaths(i)
144             If VerifyDataFolderIntegrity Then VerifyDataFolderIntegrity = False
            End If

146         i = i + 1
        Loop

    
        'If Not VerifyDataFolderIntegrity Then MsgBox sOutput, vbCritical, "OASIS folder structure corrupted"
    
        '<EhFooter>
        Exit Function

VerifyDataFolderIntegrity_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLogin.VerifyDataFolderIntegrity " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdReloadProfile_Click()
        '<EhHeader>
        On Error GoTo cmdReloadProfile_Click_Err
        '</EhHeader>

100     Me.Hide
        Dim Answer As VbMsgBoxResult
        Dim RSUpdater As ADODB.Recordset
        Dim CN As ADODB.Connection
    
102     Answer = MsgBox("WARNING! You are about to reset your profile settings?" & vbCrLf & "All your current settings will be reloaded from the server." & vbCrLf & "This may take some time and an internet connection is required. Do you want to proceed?", vbYesNo, "OASIS Maintainence")

104     If Answer = vbYes Then
 
106         Set RSUpdater = New ADODB.Recordset
108         Set CN = New ADODB.Connection
110         CN.Open GetConnectionString(g_sAppPath & "\data\db\OasisClient.mdb")
            
112         With RSUpdater
            
114             .Open "SELECT * FROM AppSettings", CN, adOpenDynamic, adLockBatchOptimistic
116             .Find "SettingName = 'ProfileSettings'"

118             If .EOF Then
120                 .AddNew
122                 .Fields("SettingName").value = "ProfileSettings"
124                 .Fields("SettingValue2").value = "0"
                End If

126             .Fields("SettingValue1").value = "0"
128             .Fields("SettingValue3").value = "0"
130             .Fields("SettingValue4").value = "0"
132             .Fields("SettingValue5").value = "0"
134             .Fields("SettingValue6").value = "0"
136             .Fields("SettingValue7").value = "0"
138             .Fields("SettingValue8").value = "0"
140             .Fields("SettingValue9").value = "0"
142             .Fields("SettingValue10").value = "0"
144             .UpdateBatch adAffectCurrent

146             SafeMoveFirst RSUpdater
148             .Find "SettingName = 'MapProjectDef'"

150             If .EOF Then
152                 .AddNew
154                 .Fields("SettingName").value = "MapProjectDef"
                End If
                
156             .Fields("SettingValue1").value = 0
158             .UpdateBatch adAffectCurrent
                   

174             .Close
                'End If

            End With
            
            'On Error Resume Next
            CN.Execute "delete from Incidents_ChartSettings"
            CN.Execute "delete from GeoBookMarks"
            CN.Execute "delete from GeoBookMarksCategories"
            CN.Execute "delete from SynchHistory where stablename = 'GeoBookMarks' or stablename  = 'GeoBookMarksCategories'"
            CN.Execute "delete from SynchHistory where stablename = 'Incidents_ChartSettings' or stablename = 'iMMAPIncidents_ChartSettings'"
            CN.Execute "delete from SynchHistoryOverview where stablename = 'Incidents_ChartSettings' or stablename = 'iMMAPIncidents_ChartSettings'"
            CN.Execute "delete from SynchHistoryOverview where stablename = 'GeoBookMarks' or stablename  = 'GeoBookMarksCategories'"
            
        
176         CN.Close
178         Set CN = Nothing
180         Set RSUpdater = Nothing
        
182         MsgBox "Profile Settings have been reset.  Next time you login online they will be updated.", vbInformation
        
        End If

184     Me.Show

        '<EhFooter>
        Exit Sub

cmdReloadProfile_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLogin.cmdReloadProfile_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdRestoreSQL_Click()
    Dim vbResponse As VbMsgBoxResult
    vbResponse = MsgBox("Are you sure you want to restore the MSSQL server database to the backup?", vbYesNo)

    If vbResponse = vbYes Then
        Me.Visible = False
        MSSQL_RestoreFromBackup
        Me.Visible = True
    End If

End Sub

Private Sub cmdUpdatePassword_Click()
        '<EhHeader>
        On Error GoTo cmdUpdatePassword_Click_Err
        '</EhHeader>
        Dim sResult As String
        Dim CN As ADODB.Connection
    
100     Me.Hide
102     'sResult = OpenSilentHttpCommsResponse("http://" & g_sAppServerPath & "/oasis4.asp?changepwd=" & txtUserName.Text & "&old=" & dxMaskEditOldPass & "&new=" & dxMaskEditNewPass1, True)
        sResult = OpenServerResponseCompressed("http://" & g_sAppServerPath & "/oasis4.asp", "changepwd", txtUserName.Text & "|||" & COnvertToMD5Pass(dxMaskEditOldPass) & "|||" & COnvertToMD5Pass(dxMaskEditNewPass1))

104     MsgBox sResult, vbInformation, "Password change"

106     If sResult = "done" Then
108         Set CN = New ADODB.Connection
110         CN.Open GetConnectionString(g_sAppPath & "\data\db\OasisClient.mdb")
112         CN.Execute "UPDATE [Personnell] SET [pwd] = '" & dxMaskEditNewPass1 & "' WHERE [UserName] = '" & txtUserName.Text & "'"
114         CN.Close
116         Set CN = Nothing
        End If

118     Me.Show
        '<EhFooter>
        Exit Sub

cmdUpdatePassword_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLogin.cmdUpdatePassword_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComServer_Click()
'    txtPassword.Text = ""
   ' txtUserName.Text = ""
  '  ComServer.toolTipText = "Current Server: " & ComServer.List(ComServer.ListIndex)
    
  '  If m_bInitReady Then KillerOnTheLoose True
End Sub

Private Sub ComServer_DropDown()
    AutoSizeBox Me.ComServer, False
End Sub



Private Sub dxMaskEditNewPass1_KeyUp(KeyCode As Integer, Shift As Integer)
CheckPasswordStuff
End Sub



Private Sub CheckPasswordStuff()
    If Len(txtUserName) > 0 And Len(dxMaskEditNewPass1) > 0 And Len(dxMaskEditNewPass2) > 0 And Len(dxMaskEditOldPass) > 0 And dxMaskEditNewPass1 = dxMaskEditNewPass2 And Not dxMaskEditNewPass1 = dxMaskEditOldPass Then
        cmdUpdatePassword.Enabled = True
    Else
    cmdUpdatePassword.Enabled = False
    End If
End Sub

Private Sub dxMaskEditNewPass2_KeyUp(KeyCode As Integer, Shift As Integer)
CheckPasswordStuff
End Sub



Private Sub dxMaskEditOldPass_KeyUp(KeyCode As Integer, Shift As Integer)
CheckPasswordStuff
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        Dim mFileSysObj As New FileSystemObject
        Dim a_strArgs() As String
        Dim CN As ADODB.Connection
        Dim oRS As ADODB.Recordset
        
        Dim i As Integer
        lStartWidth = Me.Width
        'Me.Width = Frame1.Width + 120
        
        ComServer.ListIndex = 0
        c1Tab.TabIndex = 0
        GetConnectionString ""
        If MSSQL_CheckIfInstalled Or g_sManualSQLServerPath <> "localhost" Then OptMicrosoftSQL.Enabled = True
        If bSQLServerInUse Then
            OptMicrosoftSQL.value = True
        Else
            OptMicrosoftAccess.value = True
        End If

        '        Dim RsLog As New ADODB.Recordset
        '
        '        RsLog.Open "SELECT * FROM Personnell WHERE Personnell_ID = 2", m_Cnn, adOpenDynamic, adLockOptimistic
        '
        '        If Not RsLog.EOF And Not RsLog.Bof Then
        '            If RsLog.Fields.Item("OrganisationID").Value = 1 Then
        '                chkRememberUser.Value = vbChecked
        '                txtUserName.Text = RsLog.Fields.Item("UserName").Value
        '            End If
        '        End If
        '        'm_Cnn.Execute "UPDATE Personnell SET UserName = '" & txtUserName.Text & "', pwd = '" & txtPassword.Text & "' WHERE Personnell_ID = 2"
        '        RsLog.Close
        '        Set RsLog = Nothing
        
        'keith: petri, i think you have not checked in all code
        'chkUseAutomatic.Value = IIf(CBool(g_udtSynchUpdateOptions.SynchMode), vbChecked, vbUnchecked)
        
100     If Not FileExists(g_sAppPath & "\data\db\OasisClient.mdb") Then
102         MsgBox "The OASIS database which should be located at location '" & g_sAppPath & "\data\db\OasisClient.mdb" & "' is not available.  Please contact your OASIS Administrator for assistance.", vbCritical, "Database not found"
104         End
        End If
        
106     If Not VerifyDataFolderIntegrity Then
            'End
        End If
        
108   '  g_ServerConnTimeoutSeconds = 60
110   '  g_ServerConnNoOfRetries = 3
        
        'experimental
112     m_bPrevLogin = True
        
114     'Me.caption = "OASIS Login Version: " & App.major & "." & App.minor & "." & App.Revision & " " & App.Comments
        Me.caption = App.Title & " [" & App.major & "." & App.minor & "." & App.Revision & "]"

116     Set m_HK = New cRegHotKey
118     m_HK.Attach Me.hwnd
120     m_HK.RegisterKey "Killer", vbKeyEscape, MOD_ALT + MOD_CONTROL
122     m_HK.RegisterKey "LoginD", vbKeyD, MOD_ALT + MOD_CONTROL
124     m_HK.RegisterKey "NewDB", vbKeyN, MOD_ALT + MOD_CONTROL
126     m_HK.RegisterKey "Sync", vbKeyS, MOD_ALT + MOD_CONTROL
128     m_HK.RegisterKey "ConnSpeed", vbKeyT, MOD_ALT + MOD_CONTROL
        m_HK.RegisterKey "Debug", vbKeyA, MOD_ALT + MOD_CONTROL
                
130     MainToolTip
    
132     m_Login.CreateBalloon cmdOK, "It might take some time to read all settings and data..." & vbCrLf & "Please, be patient...", "OASIS Loading Procedures...", 1

134     m_Login.ForeColor = &HEFEFEF
136     m_Login.BackColor = &HC08000

'138     Set g_PictureDialogLarge = Me.PictureDialogLarge.Picture
'140     Set g_PictureDialogSmall = Me.PictureDialogSmall.Picture
'142     Set g_PictureDialogLogo = Me.PictureDialogLogo.Picture

        'Put Icon in the SysTray
144     Call SystrayOn(Me, App.Title & " [" & App.major & "." & App.minor & "." & App.Revision & "]") ' App.major & "." & App.minor & "." & App.Revision & " " & App.Comments)
  
        Dim sCon As String
        Dim bCon As Boolean
    
146     bCon = CheckInternConnection(sCon)
    
148     If bCon Then
150         sCon = "Internet Available for Synchronisation: " & sCon
        Else
152         sCon = "Internet Not Available for Synchronisation: " & sCon
        End If
    
154     m_bInternetOK = bCon
    
156     PopupBalloon Me, "OASIS Client ready for login..." & vbCrLf & sCon, App.Title & " [" & App.major & "." & App.minor & "." & App.Revision & "]"
        
158     m_Login.Active = True
        
        On Error Resume Next
        
        Dim Process As Variant

160     For Each Process In GetObject("winmgmts:").ExecQuery("select name from Win32_Process where name='OASIS_SynchNG.exe'")
162         Process.Terminate (0)
        Next
        
        On Error GoTo Form_Load_Err
        
164     If Len(Command$) > 1 Then
            If InStr(Command$, "a-") > 0 Then
166             MainCommand
168             cmdOK_Click
                m_bInitReady = True
                Exit Sub
            Else
                a_strArgs = Split(Command$, """")

                For i = LBound(a_strArgs) To UBound(a_strArgs)

                    If InStr(a_strArgs(i + 1), ">>") Then
                        g_sAppServerPath = "http://" & Mid$(a_strArgs(i + 1), InStr(a_strArgs(i + 1), ">>") + 2)
                        cmdRestoreSQL.Enabled = IIf(FileExists(g_sAppPath & "\data\db\" & Replace(g_sAppServerPath, "/", "-") & ".bak"), True, False)
                        Exit For
                    End If

                Next
                
            End If
        End If
          
170     If Not mFileSysObj.FileExists(g_sAppPath & "\data\user\Sessions\start.dat") Then
    
172         mFileSysObj.CreateTextFile g_sAppPath & "\data\user\Sessions\start.dat", True
174         SaveStartSettings True
        
        End If
            
        Dim oIni As New clIniReader
        Dim sServers() As String
        
        i = 0
        
        CreateNewINI g_sAppPath & "\data\user\Sessions\sup.ini", oIni
    
        oIni.Path = g_sAppPath & "\data\user\Sessions\sup.ini"
        oIni.Section = "default"
        oIni.Key = "Servers"
            
        If Len(oIni.value) > 0 Then
            On Error Resume Next
                
            ComServer.Clear
                
            sServers = Split(oIni.value, ",")
                
            For i = LBound(sServers) To UBound(sServers)
                ComServer.AddItem sServers(i)
            Next
                
            If UBound(sServers) > 0 Then
                ComServer.Enabled = True
            End If
                
            oIni.Key = "DefServer"
            FindIndexStrEx ComServer, oIni.value
        End If
        
        oIni.Section = "default"
        oIni.Key = "Database"
        ComDatabase.Text = oIni.value

        If Len(ComDatabase.Text) < 2 Then ComDatabase.Text = "Microsoft Access Database"
        
        oIni.Key = "DefServer"
        ComServer.Text = Replace(oIni.value, "http://", "")

        If Len(ComServer.Text) < 2 Then ComServer.Text = "atlantis.oasiswebservice.org"
        
        oIni.Key = "RememberMe"
        chkRememberUser.value = IIf(oIni.value = "true", vbChecked, vbUnchecked)

        If chkRememberUser.value = vbChecked Then
           
            Set CN = New ADODB.Connection
            CN.Open GetConnectionString(g_sAppPath & "\data\db\OasisClient.mdb")
            Set oRS = New ADODB.Recordset
            oRS.Open "SELECT top 1 [UserName] from [Personnell]", CN, adOpenDynamic, adLockReadOnly

            If Not oRS.EOF Then
                If Not oRS.Fields(0).value = "bart" Then txtUserName.Text = oRS.Fields(0).value
                DoEvents
            End If

            oRS.Close
            CN.Close
            Set CN = Nothing
            Set oRS = Nothing
            
        End If
        
        m_bInitReady = True
        sInitialMSACCESSServer = ComServer.Text
        g_sAppServerPath = "http://" & ComServer.Text
        cmdRestoreSQL.Enabled = IIf(FileExists(g_sAppPath & "\data\db\" & Replace(g_sAppServerPath, "/", "-") & ".bak"), True, False)
        
    If ComServer.Text = "atlantis.oasiswebservice.org" Then
        txtUserName = "demo"
        txtPassword = "demo"
      '  txtPassword.Enabled = False
      '  txtUserName.Enabled = False
        
    Else
        txtPassword.Enabled = True
        txtUserName.Enabled = True
    End If
        
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub initProcedure()
'        '<EhHeader>
'        On Error GoTo initProcedure_Err
'        '</EhHeader>
'        Dim Phrase As String, Position As Integer, Asc1 As Long, Char1 As String
'        Dim txt As String
'        Dim txt1 As String
'        Dim mFile As File, mFileSysObj As New FileSystemObject, mTxtStream As TextStream
'        Dim outputText As String
'        Dim sTemp() As String
'        Dim sUserString As String
'        Dim sPasswordString As String
'
'        Exit Sub
        
'100     Set mFile = mFileSysObj.GetFile(g_sAppPath & "\data\user\Sessions\start.dat")
'102     Set mTxtStream = mFile.OpenAsTextStream(ForReading)
'104     txt = mTxtStream.ReadLine
'106     Phrase = txt
'
'108     For Position = Len(Phrase) To 1 Step -1
'110         Char1 = Mid$(Phrase, Position, 1)
'
'112         Asc1 = Asc(Char1)
'
'114         Asc1 = (((Asc1 * Asc1) / 2) / 2)
'116         Asc1 = Sqr(Asc1)
'
'118         Char1 = Chr$(Asc1)
'
'120         outputText = outputText & Char1
'        Next
'
'122     DebugPrint outputText
'
'124     m_LoginParams = Split(outputText, vbCrLf)
'
'
''        m_LoginParams(1) = "http://" & ComServer.List(ComServer.ListIndex) & "/oasis4.asp"
'
'126     If m_LoginParams(0) = "prev=2" Then
'            '            g_bDemoLogin = True
'128         m_LoginParams(0) = "prev=1"
'        End If
'
'130     If m_LoginParams(0) = "prev=1" Or g_bDemoLogin Then
'
'132         sUserString = Replace(m_LoginParams(2), "user = ", "")
'134         sPasswordString = Replace(m_LoginParams(3), "pass = ", "")
'
'136         If sUserString = "password" Then
'138             g_ClientDBPassword = sPasswordString
'            Else
'140             g_ClientDBPassword = "none"
'            End If
'
'142         m_bPrevLogin = True
'144         Set m_Cnn = New adodb.Connection
'146         m_Cnn.CursorLocation = g_sGlobalCursorLocation 'This was adUseServer
'
'            'm_Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\Oasisclient.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
'148         m_Cnn.Open GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")
'        Else
'150         m_bPrevLogin = False
'        End If
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' SettingValue5 IS USED FOR DYNAMIC CONTENT
'' for more info: http://oasis.comindwork.com/web2.aspx/OASIS/CASE5
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''        Dim rslang As New ADODB.Recordset
''
''152     If m_Cnn.State = adStateClosed Then
''            'm_Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\Oasisclient.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
''154         m_Cnn.Open GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")
''        End If
''
''156     rslang.Open "SELECT SettingValue5 FROM AppSettings WHERE SettingName = 'ProfileSettings'", m_Cnn, adOpenDynamic, adLockReadOnly
''
''158     If Not rslang.EOF And Not rslang.BOF Then
''160         If Not rslang.Fields.Item("SettingValue5").Value = vbNull Then
''162             g_sLanguage = rslang.Fields.Item("SettingValue5").Value
''
''164             If UCase$(g_sLanguage) = "DEFAULT" Then g_sLanguage = ""
''
''            End If
''        End If
''
''166     LoadLanguage Me.Name, g_sLanguage, m_Cnn
''
''168     rslang.Close
''170     Set rslang = Nothing
'
'172     sTemp = Split(m_LoginParams(4), "=")
'174     g_sAppServerPath = sTemp(1)
'        '<EhFooter>
'        Exit Sub
'
'initProcedure_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmLogin.initProcedure " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Function InitialiseConnection()
        '<EhHeader>
        On Error GoTo InitialiseConnection_Err
        '</EhHeader>

        Dim Phrase As String, Position As Integer, Asc1 As Long, Char1 As String
        Dim Txt As String
        Dim txt1 As String
        Dim mFile As File, mFileSysObj As New FileSystemObject, mTxtStream As TextStream
        Dim outputText As String
        Dim sTemp() As String
        Dim bPreppedMSSQLdb As Boolean
        bPreppedMSSQLdb = False
    
        Dim sUserString As String
        Dim sPasswordString As String
    
100     InitialiseConnection = False
102     'Me.caption = "OASIS Login Version: " & App.Title & "(" & App.major & "." & App.minor & "." & App.Revision & ")"
        Me.caption = App.Title & " [" & App.major & "." & App.minor & "." & App.Revision & "]"
        
104     If m_Cnn.State = adStateOpen Then
    
106         InitialiseConnection = True
            '  GetServerCommsSettings g_sAppPath & "\data\db\OasisClient.mdb"
        
        Else
    
108         If Not mFileSysObj.FileExists(g_sAppPath & "\data\user\Sessions\start.dat") Then
    
110             mFileSysObj.CreateTextFile g_sAppPath & "\data\user\Sessions\start.dat", True
112             SaveStartSettings True
                'SaveStartSettings
            Else
    
114             Set mFile = mFileSysObj.GetFile(g_sAppPath & "\data\user\Sessions\start.dat")
116             Set mTxtStream = mFile.OpenAsTextStream(ForReading)
118             Txt = mTxtStream.ReadLine
120             Phrase = Txt

122             For Position = Len(Phrase) To 1 Step -1
124                 Char1 = Mid$(Phrase, Position, 1)
        
126                 Asc1 = Asc(Char1)

128                 Asc1 = (((Asc1 * Asc1) / 2) / 2)
130                 Asc1 = Sqr(Asc1)

132                 Char1 = Chr$(Asc1)
                
134                 outputText = outputText & Char1
                Next
    
136             'DebugPrint outputText

138             m_LoginParams = Split(outputText, vbCrLf)
        
140             If m_LoginParams(0) = "prev=2" Then
                    '            g_bDemoLogin = True
142                 m_LoginParams(0) = "prev=1"
                End If
        
144             sTemp = Split(m_LoginParams(4), "=")

                If g_sAppServerPath = "" Then
146                 g_sAppServerPath = sTemp(1)

                    If Not g_sAppServerPath = "http://" & ComServer.List(ComServer.ListIndex) Then
                        g_sAppServerPath = "http://" & ComServer.List(ComServer.ListIndex)
                    End If

                    cmdRestoreSQL.Enabled = IIf(FileExists(g_sAppPath & "\data\db\" & Replace(g_sAppServerPath, "/", "-") & ".bak"), True, False)
                End If
                
148             If m_LoginParams(0) = "prev=1" Or g_bDemoLogin Then
        
150                 sUserString = Replace(m_LoginParams(2), "user = ", "")
152                 sPasswordString = Replace(m_LoginParams(3), "pass = ", "")
        
154                 '  If sUserString = "password" Then
156                 '   g_ClientDBPassword = sPasswordString
                    '  Else
158                 g_ClientDBPassword = "none"
                    ' End If
        
160                 m_bPrevLogin = True
162                 Set m_Cnn = New ADODB.Connection
164
                    
166

                    If InStr(GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb"), "SQLNCLI10") > 0 Then
        
                        On Error GoTo databaseconnerrorMSSQL
                        m_Cnn.CursorLocation = g_sGlobalCursorLocation
                        m_Cnn.Open GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")
            
                        On Error GoTo databaseconnerror

                        If bPreppedMSSQLdb Then m_Cnn.Open GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")
                    Else
                        m_Cnn.CursorLocation = g_sGlobalCursorLocation 'This was adUseServer
                        On Error GoTo databaseconnerror
                        m_Cnn.Open GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")
                    End If
        
                    CheckIfDebugEnhancedEnabled m_Cnn

                    On Error Resume Next
                
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    ' SettingValue5 IS USED FOR DYNAMIC CONTENT
                    ' for more info: http://oasis.comindwork.com/web2.aspx/OASIS/CASE5
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '                    Dim rslang As New ADODB.Recordset
                    '
                    '168                 If m_Cnn.State = adStateClosed Then
                    '                        'm_Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\Oasisclient.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
                    '170                     m_Cnn.Open GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")
                    '                    End If
                    '
                    '172                 rslang.Open "SELECT SettingValue5 FROM AppSettings WHERE SettingName = 'ProfileSettings'", m_Cnn, adOpenDynamic, adLockReadOnly
                    '
                    '174                 If Not rslang.EOF And Not rslang.BOF Then
                    '176                     If Not rslang.Fields.Item("SettingValue5").Value = vbNull Then
                    '178                         g_sLanguage = rslang.Fields.Item("SettingValue5").Value
                    '
                    '180                         If UCase$(g_sLanguage) = "DEFAULT" Then g_sLanguage = ""
                    '
                    '                        End If
                    '                    End If
                    '
                    '182                 LoadLanguage Me.Name, g_sLanguage, m_Cnn
                    '
                    '184                 rslang.Close
                    '186                 Set rslang = Nothing
                
                Else
188                 m_bPrevLogin = False
                End If
    
190             InitialiseConnection = True
                '   GetServerCommsSettings g_sAppPath & "\data\db\OasisClient.mdb"
            End If
    
        End If

        Exit Function
    
databaseconnerror:
        
        MsgBox "Failed to connect to the OASIS database", vbExclamation
192     InitialiseConnection = False
        Exit Function

databaseconnerrorMSSQL:

        If Trim(Err.Description) = "SQL Server Network Interfaces: Error Locating Server/Instance Specified [xFFFFFFFF]." Then
            MsgBox g_sManualSQLServerPath & "/OASISSQL MSSQL instance not found!", vbInformation
            bPreppedMSSQLdb = False
        Else
        
            MsgBox "OASIS will now configure your database for the first time.....", vbInformation
            MSSQL_PrepareNewDatabase
            bPreppedMSSQLdb = True
            Resume Next
        End If
    
        '<EhFooter>
        Exit Function

InitialiseConnection_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.InitialiseConnection " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Private Sub cmdCancel_Click()
    SystrayOff Me
    OK = False
    Unload Me 'Me.Hide
End Sub

Private Function GetPassword() As String
If bOASISv3Login Then
    GetPassword = txtPassword.Text
    
    Else
    
    Dim oMD5 As New clsMD5
    GetPassword = UCase(oMD5.MD5(StrConv(txtPassword.Text, vbUnicode)))
    End If
    

End Function

Private Function COnvertToMD5Pass(sString) As String

    Dim oMD5 As New clsMD5
    COnvertToMD5Pass = UCase(oMD5.MD5(StrConv(sString, vbUnicode)))

End Function

Private Sub SetPassword(sString)

    txtPassword.Text = sString

End Sub

Private Sub cmdOK_Click()
        '<EhHeader>
        On Error GoTo cmdOk_Click_Err
        '</EhHeader>
        
        Dim RS As New ADODB.Recordset
        Dim iGISGridVer As Integer
        Dim idynContentVer As Integer
        Dim iMapsVer As Integer
        Dim iDynamDataDefs As Integer
        Dim iSynchTables As Integer
        Dim iProjectDef As Integer
        Dim iPrintTpl As Integer
        Dim iThemes As Integer
        Dim iWebTiles As Integer
        Dim iSQLLayers As Integer
        Dim sResult As String
        Dim sString As String
        Dim bOldLoginSuccess As Boolean
        Dim RSUpdater As ADODB.Recordset
        Dim sLoginMess As String
        Dim sTitle As String
        
        Dim sCon As String
        m_bInternetOK = CheckInternConnection(sCon)
        
        cmdCancel.Enabled = False
        
        If right$(txtUserName.Text, 8) = "@oasisv3" Then
        
            bOASISv3Login = True
            txtUserName.Text = left$(txtUserName.Text, Len(txtUserName.Text) - 8)
        
        Else
        
            bOASISv3Login = False
        
        End If
        
        sTitle = "Login"
        sLoginMess = "The password did not match the previous users login!" & vbCrLf & "Would you like to login as a different user (network connection required)?"
        g_sAppServerPath = "http://" & ComServer.Text
        cmdRestoreSQL.Enabled = IIf(FileExists(g_sAppPath & "\data\db\" & Replace(g_sAppServerPath, "/", "-") & ".bak"), True, False)
        
100     If Not g_bOnlineCheckedAtLogin Then
102         g_bOnlineCheckedAtLogin = IIf(chkWorkOnline.value = vbChecked, True, False)
        End If

        'keith: petri, i think you have not checked in all code
        'g_udtSynchUpdateOptions.SynchMode = IIf(chkUseAutomatic.Value = vbChecked, True, False)

        Dim OASISSynchFolderImporter As New clSynchFolderImporter
104     OASISSynchFolderImporter.ScanAndProcessSynchFolder g_sAppPath & "\Data\Sync\import", g_sAppPath & "\Data\Db\OASISClient.mdb"
106     Set OASISSynchFolderImporter = Nothing

108     FormOnTopEx frmLogin.hwnd, False
110     cmdOK.Enabled = False
112     txtUserName.Enabled = False
114     txtPassword.Enabled = False
116     g_sUserName = txtUserName.Text
118     g_sUserPass = GetPassword  ' txtPassword.Text
            
120     If m_bChangeDefaultUser Then
122         m_bPrevLogin = False

            'm_Cnn.Execute "UPDATE AppSettings SET SettingValue1 = ['0'],SettingValue3 = ['0'],SettingValue4 = ['0'],SettingValue5 = ['0'],SettingValue6 = ['0'],SettingValue7 = ['0'],SettingValue8 = ['0'],SettingValue9 = ['0'],SettingValue10 = ['0'] WHERE SettingName = 'ProfileSettings'" '"UPDATE AppSettings SET SettingValue1 = 0 WHERE SettingName = 'ProfileSettings'"
            
124         Set RSUpdater = New ADODB.Recordset

126         With RSUpdater
            
128             .Open "SELECT * FROM AppSettings", m_Cnn, adOpenDynamic, adLockBatchOptimistic
130             .Find "SettingName = 'ProfileSettings'"

132             If .EOF Then
134                 .AddNew
136                 .Fields("SettingName").value = "ProfileSettings"
138                 .Fields("SettingValue2").value = "0"
                End If

140             .Fields("SettingValue1").value = "0"
142             .Fields("SettingValue3").value = "0"
144             .Fields("SettingValue4").value = "0"
146             .Fields("SettingValue5").value = "0"
148             .Fields("SettingValue6").value = "0"
150             .Fields("SettingValue7").value = "0"
152             .Fields("SettingValue8").value = "0"
154             .Fields("SettingValue9").value = "0"
156             .Fields("SettingValue10").value = "0"
158             .UpdateBatch adAffectCurrent

                SafeMoveFirst RSUpdater
                .Find "SettingName = 'MapProjectDef'"

                If .EOF Then
                    .AddNew
                    .Fields("SettingName").value = "MapProjectDef"
                End If
                
                .Fields("SettingValue1").value = 0
                .UpdateBatch adAffectCurrent

180             .Close
                'End If

            End With

182         Set RSUpdater = Nothing

184         m_Cnn.Execute "DELETE FROM [Incidents_ChartSettings]"
186         m_Cnn.Execute "DELETE FROM [SynchHistory] WHERE [sTableName] = 'Incidents_ChartSettings'"
            
        End If
    
188     If m_bPrevLogin Then
            
190         If InitialiseConnection Then
            
192             If g_bDemoLogin Then
                    'm_Cnn.Execute "UPDATE Personnell SET UserName = '" & txtUserName.Text & "', pwd = '" & txtPassword.Text & "' WHERE Personnell_ID = 2"

194                 Set RSUpdater = New ADODB.Recordset

196                 With RSUpdater
            
198                     .Open "SELECT * FROM Personnell", m_Cnn, adOpenDynamic, adLockBatchOptimistic
200                     .Find "Personnell_ID = 2"

202                     If Not .EOF Then
204                         .Fields("UserName").value = txtUserName.Text
206                         .Fields("pwd").value = GetPassword  'txtPassword.Text
208                         .UpdateBatch adAffectCurrent
210                         .Close
                        End If

                    End With

212                 Set RSUpdater = Nothing

214                 RS.Open "SELECT UserName, Personnell_ID FROM Personnell WHERE [pwd] ='" & GetPassword & "' AND UserName ='" & txtUserName.Text & "'", m_Cnn, adOpenForwardOnly, adLockReadOnly
                Else
216                 RS.Open "SELECT UserName, Personnell_ID FROM Personnell WHERE [pwd] ='" & GetPassword & "' AND UserName ='" & txtUserName.Text & "'", m_Cnn, adOpenForwardOnly, adLockReadOnly
                End If

            Else
                GoTo nodatabase
            End If

218         bOldLoginSuccess = False
                        
220         If RS.State = adStateOpen Then

222             If Not RS.EOF And Not RS.Bof Then
                
224                 bOldLoginSuccess = True

                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'Successfully accessed the client db and found the same login params
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
226                 g_CurrentUserID = RS.Fields("Personnell_ID").value
228                 OK = True

                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'Update profile checked
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
230                 If chkWorkOnline.value = vbChecked Then

                        frmSplash.Show vbModeless

                        DoEvents

232                     Set RS = New ADODB.Recordset
234                     RS.Open "SELECT * FROM AppSettings", m_Cnn, adOpenDynamic, adLockBatchOptimistic

                        On Error Resume Next
236                     SafeMoveFirst RS
238                     RS.Filter = "SettingName = 'MapProjectDef'"
                    
240                     If RS.EOF Then

242                         RS.AddNew
244                         RS.Fields("SettingName").value = "MapProjectDef"
246                         RS.Fields("SettingValue2").value = "0"
248                         RS.Fields("SettingValue1").value = "0"
250                         RS.Fields("SettingValue3").value = "0"
252                         RS.Fields("SettingValue4").value = "0"
254                         RS.Fields("SettingValue5").value = "0"
256                         RS.Fields("SettingValue6").value = "0"
258                         RS.Fields("SettingValue7").value = "0"
260                         RS.Fields("SettingValue8").value = "0"
262                         RS.Fields("SettingValue9").value = "0"
264                         RS.Fields("SettingValue10").value = "0"
266                         RS.UpdateBatch adAffectCurrent

                        End If

268                     If Not IsNull(RS.Fields.Item("SettingValue1").value) Then
270                         iProjectDef = RS.Fields.Item("SettingValue1").value
                        End If

308                     RS.Filter = adFilterNone
310                     SafeMoveFirst RS
312                     RS.Filter = "SettingName = 'ProfileSettings'"
                    
314                     If RS.EOF Then

316                         RS.AddNew
318                         RS.Fields("SettingName").value = "ProfileSettings"
320                         RS.Fields("SettingValue2").value = "0"
322                         RS.Fields("SettingValue1").value = "0"
324                         RS.Fields("SettingValue3").value = "0"
326                         RS.Fields("SettingValue4").value = "0"
328                         RS.Fields("SettingValue5").value = "0"
330                         RS.Fields("SettingValue6").value = "0"
332                         RS.Fields("SettingValue7").value = "0"
334                         RS.Fields("SettingValue8").value = "0"
336                         RS.Fields("SettingValue9").value = "0"
338                         RS.Fields("SettingValue10").value = "0"
340                         RS.UpdateBatch adAffectCurrent

                        End If
                        
342                     If Not IsNull(RS.Fields.Item("SettingValue4").value) Then
344                         iGISGridVer = CInt(RS.Fields.Item("SettingValue4").value)
                        End If
                        
346                     If Not IsNull(RS.Fields.Item("SettingValue5").value) Then
348                         idynContentVer = CInt(RS.Fields.Item("SettingValue5").value)
                        End If
                        
350                     If Not IsNull(RS.Fields.Item("SettingValue6").value) Then
352                         iMapsVer = CInt(RS.Fields.Item("SettingValue6").value)
                        End If
                        
354                     If Not IsNull(RS.Fields.Item("SettingValue7").value) Then
356                         iDynamDataDefs = CInt(RS.Fields.Item("SettingValue7").value)
                        End If
                        
358                     If Not IsNull(RS.Fields.Item("SettingValue8").value) Then
360                         iSynchTables = CInt(RS.Fields.Item("SettingValue8").value)
                        End If
                        
362                     If Not IsNull(RS.Fields.Item("SettingValue9").value) Then
364                         iPrintTpl = CInt(RS.Fields.Item("SettingValue9").value)
                        End If
                        
366                     If Not IsNull(RS.Fields.Item("SettingValue10").value) Then
368                         iThemes = CInt(RS.Fields.Item("SettingValue10").value)
                        End If
                                      
370                     If Not IsNull(RS.Fields.Item("SettingValue3").value) Then
372                         iSQLLayers = CInt(RS.Fields.Item("SettingValue3").value)
                        End If
                        
                        If Not IsNull(RS.Fields.Item("SettingValue3").value) Then
                            iWebTiles = CInt(RS.Fields.Item("SettingValue3").value)
                        End If
                        
374                     If g_udtSynchUpdateOptions.ForceZero Then
376                         CheckProfileUpdate txtUserName.Text, GetPassword, g_sAppServerPath & "/oasis4.asp", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, RS
                        Else
378                         CheckProfileUpdate txtUserName.Text, GetPassword, g_sAppServerPath & "/oasis4.asp", CInt(RS.Fields.Item("SettingValue1").value), iGISGridVer, idynContentVer, iMapsVer, iDynamDataDefs, iSynchTables, iPrintTpl, iThemes, iSQLLayers, iProjectDef, iWebTiles, RS
                        End If

                        frmSplash.Hide
                        'End If
                    End If

380                 If Not (Me.chkWorkOnline.value = vbChecked And g_sRemoteTablePrefix = "") Then
382                     SystrayOff Me
                        On Error Resume Next
384                     RS.Close
386                     Set RS = Nothing
                        'Me.Hide
                        m_HK.UnregisterKey "Killer"
                        m_HK.UnregisterKey "LoginD"
                        m_HK.UnregisterKey "Sync"
                        m_HK.UnregisterKey "NewDB"
                        m_HK.UnregisterKey "ConnSpeed"
                        m_HK.UnregisterKey "Debug"
388
                        frmSplash.Show vbModeless

                        DoEvents
                        Unload Me

                        If bOASISv3Login Then txtUserName.Text = txtUserName.Text & "@oasisv3"
                        cmdCancel.Enabled = True
                        Exit Sub
                    Else
390                     GoTo nointernet
                    End If

                Else
                    Set RS = New ADODB.Recordset
                    RS.Open "SELECT UserName, Personnell_ID FROM Personnell WHERE [pwd] ='simpson' AND UserName ='bart'", m_Cnn, adOpenForwardOnly, adLockReadOnly

                    If Not RS.EOF And Not RS.Bof Then
                        'First time login change the message
                        sLoginMess = "It seems like this is your first login." & vbCrLf & "Make sure you have a valid user name and password(network connection required). Do you want to continue?"
                        sTitle = "First time login"
                    End If
                End If
            End If
            
392         If Not bOldLoginSuccess Then
            
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Different login params detected
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

394             If MsgBox(sLoginMess, vbYesNo, sTitle) = vbYes Then
396                 m_bPrevLogin = False
398                 'GetClientDatabasePassword g_sAppServerPath 'sWebsite
                    'm_Cnn.Execute "UPDATE AppSettings SET SettingValue1 = '0',SettingValue3 = '0',SettingValue4 = '0',SettingValue5 = '0',SettingValue6 = '0',SettingValue7 = '0',SettingValue8 = '0',SettingValue9 = '0',SettingValue10 = '0' WHERE SettingName = 'ProfileSettings'" '"UPDATE AppSettings SET SettingValue1 = 0 WHERE SettingName = 'ProfileSettings'"

400                 Set RSUpdater = New ADODB.Recordset

402                 With RSUpdater
            
404                     .Open "SELECT * FROM [AppSettings]", m_Cnn, adOpenDynamic, adLockBatchOptimistic
406                     .Find "SettingName = 'ProfileSettings'"

408                     If Not .EOF Then
410                         .Fields("SettingDesc").value = ""
412                         .Fields("SettingValue1").value = "0"
414                         .Fields("SettingValue3").value = "0"
416                         .Fields("SettingValue4").value = "0"
418                         .Fields("SettingValue5").value = "0"
420                         .Fields("SettingValue6").value = "0"
422                         .Fields("SettingValue7").value = "0"
424                         .Fields("SettingValue8").value = "0"
426                         .Fields("SettingValue9").value = "0"
428                         .Fields("SettingValue10").value = "0"
430                         .UpdateBatch adAffectCurrent
432                         .Close
                        End If

                    End With

434                 Set RSUpdater = Nothing

436                 Me.chkWorkOnline.value = vbChecked

                    If bOASISv3Login Then txtUserName.Text = txtUserName.Text & "@oasisv3"
438                 cmdOK_Click
                    cmdCancel.Enabled = True
                    Exit Sub
                Else
440                 txtUserName.Enabled = True
442                 txtPassword.Enabled = True
444                 txtPassword.SetFocus
446                 txtPassword.SelStart = 0
448                 txtPassword.SelLength = Len(txtPassword.Text)
450                 cmdOK.Enabled = True

452                 FormOnTop Me
                End If

                On Error Resume Next
454             RS.Close
456             Set RS = Nothing

            End If

        Else
        
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Attempt login to OASIS Server with new user
             
            'No need to check for chkUseAutomatic since this is forced on profile update
        
458         If m_bInternetOK Then 'And chkUseAutomatic.Value = vbChecked) Then
            
460             If Not g_sAppServerPath = "http://" & ComServer.Text Then
462                 g_sAppServerPath = "http://" & ComServer.Text
                    cmdRestoreSQL.Enabled = IIf(FileExists(g_sAppPath & "\data\db\" & Replace(g_sAppServerPath, "/", "-") & ".bak"), True, False)
                End If
            
464             'sString = g_sAppServerPath & "/oasis4.asp?user=" & CheckEncrypt(txtUserName.Text) & "&" & "pwd=" & CheckEncrypt(txtPassword.Text)
466             'sResult = OpenSilentHttpCommsResponse(sString, True)
                sResult = OpenServerResponseCompressed(g_sAppServerPath & "/oasis4.asp", "user", txtUserName.Text & "|||" & GetPassword)
            
468             If Len(sResult) > 5 Then

470                 sResult = sResult & "/"

472                 If Not CheckProfileUpdate(txtUserName.Text, GetPassword, g_sAppServerPath & "/oasis4.asp", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) Then
474                     MsgBox "The OASIS profile update failed. Please try again", vbExclamation, "OASIS login failure"
476                     txtUserName.Enabled = True
478                     txtPassword.Enabled = True
480                     cmdOK.Enabled = True
482                     OK = False

                        If bOASISv3Login Then txtUserName.Text = txtUserName.Text & "@oasisv3"
                        cmdCancel.Enabled = True
                        Exit Sub
                    End If

484                 SaveStartSettings
                    'm_Cnn.Execute "UPDATE Personnell SET UserName = '" & txtUserName.Text & "', pwd = '" & txtPassword.Text & "' WHERE Personnell_ID = 2"   ' & g_CurrentUserID

486                 Set RSUpdater = New ADODB.Recordset

488                 With RSUpdater
            
490                     .Open "SELECT * FROM Personnell", m_Cnn, adOpenDynamic, adLockBatchOptimistic
492                     .Find "Personnell_ID = 2"

494                     If Not .EOF Then
496                         .Fields("UserName").value = txtUserName.Text
498                         .Fields("pwd").value = GetPassword  'txtPassword.Text
500                         .UpdateBatch adAffectCurrent
502                         .Close
                        End If

                    End With

504                 Set RSUpdater = Nothing

506                 chkWorkOnline.value = vbUnchecked
                Else
                
nointernet:
                
508                 MsgBox "Communication to the OASIS Server failed.  Please check your internet connection and try again", vbExclamation, "OASIS login failure"

nodatabase:
510                 txtUserName.Enabled = True
512                 txtPassword.Enabled = True
514                 cmdOK.Enabled = True
516                 OK = False

                    If bOASISv3Login Then txtUserName.Text = txtUserName.Text & "@oasisv3"
                    cmdCancel.Enabled = True
                    Exit Sub
                End If
                
518             m_bChangeDefaultUser = False

                If bOASISv3Login Then txtUserName.Text = txtUserName.Text & "@oasisv3"
520             cmdOK_Click

                Exit Sub

            Else

                'This is new - I added this since we should not let new users login
                'if internet connection is down
            
522             MsgBox "You cannot login as a new user when your internet connection is down"
524             txtUserName.Enabled = True
526             txtPassword.Enabled = True
528             cmdOK.Enabled = True
530             OK = False
            
            End If
        End If
        
        If bOASISv3Login Then txtUserName.Text = txtUserName.Text & "@oasisv3"
    
        '<EhFooter>
        cmdCancel.Enabled = True
        Exit Sub

cmdOk_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.cmdOK_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Function CheckProfileUpdateEx() As Boolean
    Dim oRSAppSetts As Recordset
    
    'oRSAppSetts
    
End Function

Private Function AppSettingsUpdate(CN As ADODB.Connection, _
                                   Optional ORSLocalAppSettings As ADODB.Recordset) As Boolean
        '<EhHeader>
        On Error GoTo AppSettingsUpdate_Err
        '</EhHeader>
        Dim sVals As String
        Dim j As Integer
        Dim sString As String
        Dim RSUpdater As ADODB.Recordset
        Dim bErrorAlreadyCaught As Boolean
        
100     If g_udtSynchUpdateOptions.lMethod = 0 Then
        
102         If SafeMoveFirst(m_RSProfile) Then
        
                DebugPrint "updating profilesettings......"
104             CN.Execute "delete from AppSettings WHERE SettingName <> 'ProfileSettings'"
                DebugPrint "CN.State = " & CN.State

                Do While CN.State = adStateFetching
                    DebugPrint "CN.State = " & CN.State
                Loop

106             Set RSUpdater = New ADODB.Recordset
108             RSUpdater.Open "SELECT * FROM AppSettings", CN, adOpenDynamic, adLockBatchOptimistic
            
110             Do While Not m_RSProfile.EOF

112                 With RSUpdater
                    
114                     If m_RSProfile.Fields.Item("SettingName").value = "ProfileSettings" Then
                        
116                         'This code is necesscary in order to avoid overriding
                            'of the profile settings values for dynamic data, synchtables, etc....

                            .Find "SettingName = 'ProfileSettings'"
                            
118                         If Not .EOF Then
120                             .Fields("SettingValue1").value = IIf((m_RSProfile.Fields.Item("SettingValue1").value) = "", Null, m_RSProfile.Fields.Item("SettingValue1").value)

                                DoEvents
122                             .UpdateBatch 'adAffectCurrent
                            End If
                        
                        Else
                    
124                         .AddNew
126                         .Fields("SettingDesc").value = m_RSProfile.Fields.Item("SettingDesc").value
128                         .Fields("SettingName").value = m_RSProfile.Fields.Item("SettingName").value
130                         .Fields("SettingValue1").value = IIf((m_RSProfile.Fields.Item("SettingValue1").value) = "", Null, m_RSProfile.Fields.Item("SettingValue1").value)
132                         .Fields("SettingValue2").value = IIf((m_RSProfile.Fields.Item("SettingValue2").value) = "", Null, m_RSProfile.Fields.Item("SettingValue2").value)
134                         .Fields("SettingValue3").value = IIf((m_RSProfile.Fields.Item("SettingValue3").value) = "", Null, m_RSProfile.Fields.Item("SettingValue3").value)
136                         .Fields("SettingValue4").value = IIf((m_RSProfile.Fields.Item("SettingValue4").value) = "", Null, m_RSProfile.Fields.Item("SettingValue4").value)
138                         .Fields("SettingValue5").value = IIf((m_RSProfile.Fields.Item("SettingValue5").value) = "", Null, m_RSProfile.Fields.Item("SettingValue5").value)
140                         .Fields("SettingValue6").value = IIf((m_RSProfile.Fields.Item("SettingValue6").value) = "", Null, m_RSProfile.Fields.Item("SettingValue6").value)
142                         .Fields("SettingValue7").value = IIf((m_RSProfile.Fields.Item("SettingValue7").value) = "", Null, m_RSProfile.Fields.Item("SettingValue7").value)
144                         .Fields("SettingValue8").value = IIf((m_RSProfile.Fields.Item("SettingValue8").value) = "", Null, m_RSProfile.Fields.Item("SettingValue8").value)
146                         .Fields("SettingValue9").value = IIf((m_RSProfile.Fields.Item("SettingValue9").value) = "", Null, m_RSProfile.Fields.Item("SettingValue9").value)
148                         .Fields("SettingValue10").value = IIf((m_RSProfile.Fields.Item("SettingValue10").value) = "", Null, m_RSProfile.Fields.Item("SettingValue10").value)
150                         .UpdateBatch adAffectCurrent
                        End If

                    End With
                                      
152                 m_RSProfile.MoveNext
                Loop

154             RSUpdater.Close
156             Set RSUpdater = Nothing
        
            End If
        
158     ElseIf g_udtSynchUpdateOptions.lMethod = 1 Then

160         If Not ORSLocalAppSettings Is Nothing Then
            
162             ORSLocalAppSettings.Filter = adFilterNone

164             If SafeMoveFirst(ORSLocalAppSettings) Then
            
166                 'sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & "AppSettings")
                    Set m_RSProfile = OpenServerRSCompressed(g_sAppServerPath & "/oasis4.asp", "ID", "SELECT * FROM " & g_sRemoteTablePrefix & "AppSettings")
168                 'Set m_RSProfile = OpenSilentHttpCommsRS(sString, True)

170                 If m_RSProfile.State = 0 Then
172                     AppSettingsUpdate = False
                    Else

174                     Do While Not ORSLocalAppSettings.EOF
                    
176                         CN.Execute "delete from AppSettings WHERE SettingName = '" & ORSLocalAppSettings.Fields.Item("SettingName").value & "'"
178                         m_RSProfile.Find "SettingName = '" & ORSLocalAppSettings.Fields.Item("SettingName").value & "'"
                    
180                         Set RSUpdater = New ADODB.Recordset

182                         With RSUpdater
                    
184                             .Open "SELECT * FROM AppSettings", CN, adOpenDynamic, adLockBatchOptimistic
186                             .AddNew
188                             .Fields("SettingName").value = m_RSProfile.Fields.Item("SettingName").value
190                             .Fields("SettingDesc").value = m_RSProfile.Fields.Item("SettingDesc").value
192                             .Fields("SettingValue1").value = m_RSProfile.Fields.Item("SettingValue1").value
194                             .Fields("SettingValue2").value = m_RSProfile.Fields.Item("SettingValue2").value
196                             .Fields("SettingValue3").value = m_RSProfile.Fields.Item("SettingValue3").value
198                             .Fields("SettingValue4").value = m_RSProfile.Fields.Item("SettingValue4").value
200                             .Fields("SettingValue5").value = m_RSProfile.Fields.Item("SettingValue5").value
202                             .Fields("SettingValue6").value = m_RSProfile.Fields.Item("SettingValue6").value
204                             .Fields("SettingValue7").value = m_RSProfile.Fields.Item("SettingValue7").value
206                             .Fields("SettingValue8").value = m_RSProfile.Fields.Item("SettingValue8").value
208                             .Fields("SettingValue9").value = m_RSProfile.Fields.Item("SettingValue9").value
210                             .Fields("SettingValue10").value = m_RSProfile.Fields.Item("SettingValue10").value
212                             .UpdateBatch adAffectCurrent
214                             .Close
                        
                            End With

216                         Set RSUpdater = Nothing
             
218                         ORSLocalAppSettings.MoveNext
                        Loop
                        
                    End If
                
220                 Set m_RSProfile = Nothing
                
                End If
            
            End If
        
        End If

        '<EhFooter>
        Exit Function

AppSettingsUpdate_Err:

        If bErrorAlreadyCaught Then
        
            MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.AppSettingsUpdate " & "at line " & Erl
        Else
            DebugPrint "AppSettings table access error detected.  Waiting 3 seconds and retrying sub [frmLogin.AppSettingsUpdate]"
            bErrorAlreadyCaught = True
            SleepAPI 3000
            GoTo 100
        
        End If

        '</EhFooter>
End Function

'Private Sub GISGridUpdate(CN As Adodb.Connection, _
'                          RSRemote As Adodb.Recordset, _
'                          RS As Adodb.Recordset)
'        '<EhHeader>
'        On Error GoTo GISGridUpdate_Err
'        '</EhHeader>
'         Dim j As Integer
'         Dim sString As String
'
'        sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & "GISGridTableSettings")
'100     Set RSRemote = OpenSilentHttpCommsRS(sString, True)
'
'102     If Not RSRemote.State = 0 Then
'104         Set RS = New Adodb.Recordset
'
'106         CN.Execute "delete from GISGridTableSettings"
'
'108         If Not RSRemote.EOF And Not RSRemote.BOF Then
'
'110             SafeMoveFirst RSRemote
'
'112             RS.Open "SELECT * FROM GISGridTableSettings", CN, adOpenDynamic, adLockOptimistic
'
'114             Do While Not RSRemote.EOF
'116                 RS.AddNew
'                     'Petri Changed this 1 to 0
'118                 For j = 0 To RSRemote.Fields.Count - 1
'120                     'RS.Fields.Item(j).Value = rsRemote.Fields.Item(j).Value
'                        RS.Fields.Item(j).Value = RSRemote.Fields(RS.Fields.Item(j).Name).Value
'                    Next
'
'122                 RSRemote.MoveNext
'                Loop
'
'124             RS.UpdateBatch
'126             RSRemote.Close
'128             RS.Close
'
'            End If
'
'            SynchProfileSettingWithServer "SettingValue4", g_sRemoteTablePrefix, CN
'130         Set RSRemote = Nothing
'132         Set RS = Nothing
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'GISGridUpdate_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmLogin.GISGridUpdate " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub

'Private Sub ThematicsUpdate(CN As Adodb.Connection, _
'                            RSRemote As Adodb.Recordset, _
'                            RS As Adodb.Recordset)
'        '<EhHeader>
'        On Error GoTo ThematicsUpdate_Err
'        '</EhHeader>
'        Dim j As Integer
'        Dim sString As String
'
'100     sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & "Themes")
'102     Set RSRemote = OpenSilentHttpCommsRS(sString, True)
'
'104     If Not RSRemote.State = 0 Then
'106         Set RS = New Adodb.Recordset
'
'108         CN.Execute "delete from Themes"
'
'110         If Not RSRemote.EOF And Not RSRemote.BOF Then
'
'112             SafeMoveFirst RSRemote
'114             Err.Clear
'
'116             If Not Err.number > 0 Then
'118                 RS.Open "SELECT * FROM Themes", CN, adOpenDynamic, adLockOptimistic
'
'120                 Do While Not RSRemote.EOF
'
'122                     If Not Err.number > 0 Then
'124                         RS.AddNew
'
'126                         For j = 0 To RSRemote.Fields.Count - 1
'                                If Not IsNull(RSRemote.Fields.Item(j).Value) Then
'128                                 'RS.Fields.Item(j).Value = rsRemote.Fields.Item(j).Value
'                                    RS.Fields.Item(j).Value = RSRemote.Fields(RS.Fields.Item(j).Name).Value
'                                End If
'                            Next
'
'                        End If
'
'130                     RSRemote.MoveNext
'
'                    Loop
'
'132                 RS.UpdateBatch
'                End If
'
'134             RSRemote.Close
'136             RS.Close
'
'            End If
'
'138         SynchProfileSettingWithServer "SettingValue10", g_sRemoteTablePrefix, CN
'
'140         Set RSRemote = Nothing
'142         Set RS = Nothing
'        End If
'
'144     sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & "ThemeGroups")
'146     Set RSRemote = OpenSilentHttpCommsRS(sString, True)
'
'148     If Not RSRemote.State = 0 Then
'150         Set RS = New Adodb.Recordset
'
'152         CN.Execute "delete from ThemeGroups"
'
'154         If Not RSRemote.EOF And Not RSRemote.BOF Then
'
'156            SafeMoveFirst RSRemote
'
'158             Err.Clear
'
'160             If Not Err.number > 0 Then
'162                 RS.Open "SELECT * FROM ThemeGroups", CN, adOpenDynamic, adLockOptimistic
'
'164                 Do While Not RSRemote.EOF
'
'166                     If Not Err.number > 0 Then
'168                         RS.AddNew
'
'170                         For j = 0 To RSRemote.Fields.Count - 1
'172                             'RS.Fields.Item(j).Value = rsRemote.Fields.Item(j).Value
'                                RS.Fields.Item(j).Value = RSRemote.Fields(RS.Fields.Item(j).Name).Value
'                            Next
'
'                        End If
'
'174                     RSRemote.MoveNext
'
'                    Loop
'
'176                 RS.UpdateBatch
'                End If
'
'178             RSRemote.Close
'180             RS.Close
'
'            End If
'
'182         Set RSRemote = Nothing
'184         Set RS = Nothing
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'ThematicsUpdate_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmLogin.ThematicsUpdate " & _
'               "at line " & Erl
'        'Resume Next
'        '</EhFooter>
'End Sub

'Private Sub SynchLayerUpdate(CN As Adodb.Connection, _
'                             RSRemote As Adodb.Recordset, _
'                             RS As Adodb.Recordset, _
'                             iClientSynchTables As Integer)
'        '<EhHeader>
'        On Error GoTo SynchLayerUpdate_Err
'        '</EhHeader>
'        Dim j As Integer
'        Dim RSRemoteReplicator As Adodb.Recordset
'        Dim sString As String
'        'Dim RSUpdater As ADODB.Recordset
'
'100     sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & "SynchTables")
'102     Set RSRemote = OpenSilentHttpCommsRS(sString, True)
'
'104     If Not RSRemote.State = 0 Then
'106         Set RS = New Adodb.Recordset
'
'108         If Not RSRemote.EOF And Not RSRemote.BOF Then
'
'110             SafeMoveFirst RSRemote
'
'112             CN.Execute "delete from SynchTables"
'114             RS.Open "SELECT * FROM SynchTables", CN, adOpenDynamic, adLockOptimistic
'
'116             Do While Not RSRemote.EOF
'
'118                 If Not DoesTableExist(m_Cnn.ConnectionString, RSRemote.Fields.Item("sTableName").Value) Then
'
'120                     sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & RSRemote.Fields.Item("sTableName").Value)
'122                     Set RSRemoteReplicator = OpenSilentHttpCommsRS(sString, True)
'
'124                     If Not RSRemoteReplicator Is Nothing Then ' = adStateOpen Then
'126                         CreateTable RSRemote.Fields.Item("sTableName").Value, RSRemoteReplicator, m_Cnn
'                        End If
'
'                    End If
'
'128                 RS.AddNew
'
'130                 For j = 0 To RSRemote.Fields.Count - 1
'132                     'RS.Fields.Item(j).Value = rsRemote.Fields.Item(j).Value
'                        RS.Fields.Item(j).Value = RSRemote.Fields(RS.Fields.Item(j).Name).Value
'                    Next
'
'134                 RSRemote.MoveNext
'                Loop
'
'136             RS.UpdateBatch
'138             RSRemote.Close
'140             RS.Close
'
'            End If
'        End If
'
'        On Error Resume Next
'
'142     Set RSRemote = Nothing
'144     Set RS = Nothing
'
'        'm_Cnn.Execute "UPDATE AppSettings SET SettingValue8 = '" & iClientSynchTables & "' WHERE SettingName = 'ProfileSettings'"
'
''146     Set RSUpdater = New ADODB.Recordset
''148     With RSUpdater
''
''150         .Open "SELECT * FROM AppSettings", m_Cnn, adOpenDynamic, adLockBatchOptimistic
''152         .Find "SettingName = 'ProfileSettings'"
''
''154         If Not .EOF Then
''156             .Fields("SettingValue8").Value = CStr(iClientSynchTables)
''158             .UpdateBatch adAffectCurrent
''160             .Close
''            End If
''
''        End With
''162     Set RSUpdater = Nothing
'
'        SynchProfileSettingWithServer "SettingValue8", g_sRemoteTablePrefix, m_Cnn
'
'        '<EhFooter>
'        Exit Sub
'
'SynchLayerUpdate_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmLogin.SynchLayerUpdate " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub

Private Sub GetRemoteProfile(sURL As String, _
                             iInteger As Integer, _
                             Optional sName As String, _
                             Optional sPASS As String)
        '<EhHeader>
        On Error GoTo GetRemoteProfile_Err
        '</EhHeader>
        
        Dim sString As String
        
        If sName = "" Then sName = txtUserName.Text
        If sPASS = "" Then sPASS = txtPassword.Text
        
        Set m_RSProfile = New ADODB.Recordset
        
100     Select Case g_udtSynchUpdateOptions.lMethod
            
            Case 0
                'sString = sURL & "?user=" & CheckEncrypt(sName) & "&pwd=" & CheckEncrypt(sPASS) & "&ver=" & CheckEncrypt(CStr(iInteger)) & "&sessUID=" & CheckEncrypt(GetGuid)
                'Set m_RSProfile = OpenSilentHttpCommsRS(sString, True)
                Set m_RSProfile = OpenServerRSCompressed(sURL, "user", sName & "|||" & sPASS) ' & "&ver=" & Str(iInteger) & "&sessUID=" & GetGuid)

104         Case 1

106             If Len(g_sRemoteTablePrefix) > 0 Then
                    'sString = sURL & "?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & "AppSettings WHERE SettingName = 'ProfileSettings'")
                    'Set m_RSProfile = OpenSilentHttpCommsRS(sString, True)
                    Set m_RSProfile = OpenServerRSCompressed(sURL, "ID", "SELECT * FROM " & g_sRemoteTablePrefix & "AppSettings WHERE SettingName = 'ProfileSettings'")
                Else
                    'Fails To get Update on low bandwidth... Do something
                End If

110         Case 2
        
        End Select

        '<EhFooter>
        Exit Sub

GetRemoteProfile_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.GetRemoteProfile " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub GetRemoteTablePrefix(sUser As String, _
                                 sPASS As String, _
                                 sURL As String)
        '<EhHeader>
        On Error GoTo GetRemoteTablePrefix_Err
        '</EhHeader>
        
        Dim sString As String
        
        'sString = sURL & "?ID=" & CheckEncrypt("SELECT SettingTablePrefix FROM UserGroups WHERE ID IN (SELECT UserGroupID FROM Users WHERE [user] = '" & sUser & "' AND [pwd] = '" & sPASS & "')")
        'Set mRSUGSettings = OpenSilentHttpCommsRS(sString, True)
        Set mRSUGSettings = OpenServerRSCompressed(sURL, "ID", "SELECT SettingTablePrefix FROM UserGroups WHERE ID IN (SELECT UserGroupID FROM Users WHERE [user] = '" & sUser & "' AND [pwd] = '" & sPASS & "')")
    
104     If Not mRSUGSettings.State = 0 Then

            If Not mRSUGSettings.EOF Then
106         g_sRemoteTablePrefix = mRSUGSettings.Fields.Item("SettingTablePrefix").value
            Else
            g_sRemoteTablePrefix = ""
            End If
108         mRSUGSettings.Close
110         Set mRSUGSettings = Nothing
        End If

        '<EhFooter>
        Exit Sub

GetRemoteTablePrefix_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.GetRemoteTablePrefix " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CheckProfileUpdate(sUser As String, _
                                    sPASS As String, _
                                    sURL As String, _
                                    iInteger As Integer, _
                                    iGISGrid As Integer, _
                                    IDynCont As Integer, _
                                    iMap As Integer, _
                                    iDynamDataDefs As Integer, _
                                    iSynchTables As Integer, _
                                    iMapTpl As Integer, _
                                    iThemes As Integer, _
                                    iSQLLayers As Integer, _
                                    iMapProjectDef As Integer, _
                                    iWebTiles As Integer, _
                                    Optional ORSLocalAppSettings As ADODB.Recordset) As Boolean
        '<EhHeader>
        On Error GoTo CheckProfileUpdate_Err
        '</EhHeader>
        Dim sGroupName As String
        Dim sTablePrefix As String
        Dim CN As New ADODB.Connection
        Dim RS As New ADODB.Recordset
        Dim rsRemote As ADODB.Recordset
        Dim j As Integer
        Dim iClientGISGrid As Integer
        Dim iClientDynCont As Integer
        Dim iClientMap As Integer
        Dim iClientWebTiles As Integer
        Dim iClientDynamDataDefs As Integer
        Dim iClientSynchTables As Integer
        Dim iClientMapTemplates As Integer
        Dim iClientThemes As Integer
        Dim iAppSettings As Integer
        Dim iClientSQLLayers As Integer
        Dim iClientMapProjectDef As Integer
        
100     If g_udtSynchUpdateOptions.ManualSynchronisation Then
102         CheckProfileUpdate = True
            Exit Function
        End If

104     CheckProfileUpdate = True
   
        Dim i As Integer
        
        On Error Resume Next
        
106     If g_sRemoteTablePrefix = "" Then
108         GetRemoteTablePrefix sUser, sPASS, sURL
        End If
        
110     If g_sRemoteTablePrefix = "" Then
114         CheckProfileUpdate = False
            Exit Function
        End If
        
116     GetRemoteProfile sURL, iInteger, sUser, sPASS
        
        On Error GoTo CheckProfileUpdate_Err
        
118     If m_RSProfile.State = 0 Then
120         CheckProfileUpdate = False
            Exit Function
        End If

122     If Not m_RSProfile.EOF Then
            
124         SafeMoveFirst m_RSProfile
126         m_RSProfile.Find "SettingName = 'ProfileSettings'"
128         'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\Oasisclient.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
            CN.Open GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")

            '            If Not IsNull(m_RSProfile.Fields.Item("SettingValue3").value) Then
            '                If IsNumeric(m_RSProfile.Fields.Item("SettingValue3").value) Then
            '                    iClientSQLLayers = CInt(m_RSProfile.Fields.Item("SettingValue3").value)
            '                End If
            '            End If

            If Not IsNull(m_RSProfile.Fields.Item("SettingValue3").value) Then
                If IsNumeric(m_RSProfile.Fields.Item("SettingValue3").value) Then
                    iClientWebTiles = CInt(m_RSProfile.Fields.Item("SettingValue3").value)
                End If
            End If

            'Save the versions to check later in function
130         If Not IsNull(m_RSProfile.Fields.Item("SettingValue4").value) Then
132             If IsNumeric(m_RSProfile.Fields.Item("SettingValue4").value) Then
134                 iClientGISGrid = CInt(m_RSProfile.Fields.Item("SettingValue4").value)
                End If
            End If
            
136         If Not IsNull(m_RSProfile.Fields.Item("SettingValue5").value) Then
138             If IsNumeric(m_RSProfile.Fields.Item("SettingValue5").value) Then
140                 iClientDynCont = CInt(m_RSProfile.Fields.Item("SettingValue5").value)
                End If
            End If
            
142         If Not IsNull(m_RSProfile.Fields.Item("SettingValue6").value) Then
144             If IsNumeric(m_RSProfile.Fields.Item("SettingValue6").value) Then
146                 iClientMap = CInt(m_RSProfile.Fields.Item("SettingValue6").value)
                End If
            End If
            
148         If Not IsNull(m_RSProfile.Fields.Item("SettingValue7").value) Then
150             If IsNumeric(m_RSProfile.Fields.Item("SettingValue7").value) Then
152                 iClientDynamDataDefs = CInt(m_RSProfile.Fields.Item("SettingValue7").value)
                End If
            End If
            
154         If Not IsNull(m_RSProfile.Fields.Item("SettingValue8").value) Then
156             If IsNumeric(m_RSProfile.Fields.Item("SettingValue8").value) Then
158                 iClientSynchTables = CInt(m_RSProfile.Fields.Item("SettingValue8").value)
                End If
            End If
              
160         If Not IsNull(m_RSProfile.Fields.Item("SettingValue9").value) Then
162             If IsNumeric(m_RSProfile.Fields.Item("SettingValue9").value) Then
164                 iClientMapTemplates = CInt(m_RSProfile.Fields.Item("SettingValue9").value)
                End If
            End If
            
166         If Not IsNull(m_RSProfile.Fields.Item("SettingValue10").value) Then
168             If IsNumeric(m_RSProfile.Fields.Item("SettingValue10").value) Then
170                 iClientThemes = CInt(m_RSProfile.Fields.Item("SettingValue10").value)
                End If
            End If
            
172         If Not IsNull(m_RSProfile.Fields.Item("SettingValue1").value) Then
174             If IsNumeric(m_RSProfile.Fields.Item("SettingValue1").value) Then
176                 iAppSettings = CInt(m_RSProfile.Fields.Item("SettingValue1").value)
                End If
            End If
            
            SafeMoveFirst m_RSProfile
            m_RSProfile.Find "SettingName = 'MapProjectDef'"

            If Not m_RSProfile.EOF Then
                If Not IsNull(m_RSProfile.Fields.Item("SettingValue1").value) Then
                    If IsNumeric(m_RSProfile.Fields.Item("SettingValue1").value) Then
                        iClientMapProjectDef = CInt(m_RSProfile.Fields.Item("SettingValue1").value)
                    End If
                End If
            End If
            
            SafeMoveFirst m_RSProfile
            m_RSProfile.Find "SettingName = 'ProfileSettings'"
        
            ' this was iInteger < iAppSettings, but it was changed since it could mess things up if version number in the client db were screwed up
            ' Petri: delete this comment if you are happy with this
178         If iInteger <> iAppSettings Then
180             AppSettingsUpdate CN, ORSLocalAppSettings
            End If
            
        End If

182     If g_udtSynchUpdateOptions.GISAttributeSettings Then

            'Now Check the GIS Grid version
184         If Not g_sRemoteTablePrefix = "" Then
186             If iClientGISGrid > iGISGrid Then
                    UGTableUpdate "GISGridTableSettings", CN, "SettingValue4", "ProfileSettings"
188                 'Set RSRemote = New Adodb.Recordset
190                 'GISGridUpdate CN, RSRemote, RS
                End If
            End If
        
        End If
        
192     If g_udtSynchUpdateOptions.PrintTemplates Then

            'Now Check the MapTemplates version
194         If iClientMapTemplates > iMapTpl Then
196             g_bMapTplUpdate = True
            End If
        End If
        
198     If g_udtSynchUpdateOptions.MapProducts Then

            'Checking the Maps Library
200         If iClientMap > iMap Then
202             g_bMapsUpdate = True
            End If
        End If

        'Now Check the Dynamic Content version
    
204     If g_udtSynchUpdateOptions.Feeds Then
206         If Not g_sRemoteTablePrefix = "" Then
208             If iClientDynCont > IDynCont Then
210                 g_bFeedUpdate = True
                End If
            End If
        End If
        
212     If g_udtSynchUpdateOptions.DynamDataDefs Then

            'Now check the DynamDataTable version
214         If Not g_sRemoteTablePrefix = "" Then
216             If iClientDynamDataDefs > iDynamDataDefs Then
                    g_bUpdateDynamicDataDefs = True
                    
                    UGTableUpdate "DynamicDataDefs", CN, "SettingValue7", "ProfileSettings"
                    'Set RSRemote = New Adodb.Recordset
                    'DynamDataDefsUpdate CN, RSRemote, RS
                End If
            End If
        End If
        
'        'iClientSQLLayers
'        If Not g_sRemoteTablePrefix = "" Then
'            If iClientSQLLayers > iSQLLayers Then
'
'                UGTableUpdate "ttkGISLayerSQLInProject", CN, "SettingValue3", "ProfileSettings"
'                UpdateSQLLayersInTTkGISLayerSQL CN
'                'Set RSRemote = New Adodb.Recordset
'                'SQLLayersUpdate CN, RSRemote, RS
'            End If
'        End If
        
222     If g_udtSynchUpdateOptions.Thematics Then

            'Now check the Thematics version
224         If Not g_sRemoteTablePrefix = "" Then
226             If iClientThemes > iThemes Then
                    UGTableUpdate "Themes", CN, "SettingValue10", "ProfileSettings"
                    UGTableUpdate "ThemeGroups", CN, "", ""
228                 'Set RSRemote = New Adodb.Recordset
230                 'ThematicsUpdate CN, RSRemote, RS
                End If
            End If
        End If

232     If g_udtSynchUpdateOptions.SynchLayersSettings Then

            'Now check the Synch Table version
234         If Not g_sRemoteTablePrefix = "" Then
236             If iClientSynchTables > iSynchTables Then

                    'OLD CODE
238                 'Set RSRemote = New Adodb.Recordset
240                 'SynchLayerUpdate CN, RSRemote, RS, iClientSynchTables

                End If
            End If
        End If
        
        If Not g_sRemoteTablePrefix = "" Then
            If iClientMapProjectDef > iMapProjectDef Then

                UGTableUpdate "ttkGISProjectDef", CN, "SettingValue1", "MapProjectDef"
                'Set RSRemote = New Adodb.Recordset
                'MapProjectDefUpdate CN, RSRemote, RS
            End If
            
            If iClientWebTiles > iWebTiles Then

                UGTableUpdate "WebTiles", CN, "SettingValue3", "ProfileSettings"
                'Set RSRemote = New Adodb.Recordset
                'MapProjectDefUpdate CN, RSRemote, RS
            End If
        End If

        '''''''''''''''''''''''

242     If CN.State = adStateOpen Then
244         CN.Close
        End If
        
246     Set CN = Nothing

248     CheckProfileUpdate = True

        '<EhFooter>
        Exit Function

CheckProfileUpdate_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.CheckProfileUpdate " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub ReNewAppSettings(sVals As String)
        '<EhHeader>
        On Error GoTo ReNewAppSettings_Err
        '</EhHeader>
    
        Dim sTablePrefix As String
        Dim WebSite As String
        Dim Records As ADODB.Recordset
        Dim RSAppSetting As ADODB.Recordset
        Dim sSQL As String
        Dim sString As String
    
102     Set Records = New ADODB.Recordset
104     Set RSAppSetting = New ADODB.Recordset
    
        'TODO !!!
        MsgBox "You should not get here. Who told you that?", vbExclamation
        Stop
106     WebSite = "http://www.immap.org/"
    
108     sSQL = "SELECT SettingTablePrefix FROM UserGroups WHERE ID IN (SELECT UserGroupID FROM Users WHERE [user] = '" & txtUserName.Text & "' AND [pwd] = '" & GetPassword & "')"

        'sString = WebSite & "oasis4.asp?ID=" & CheckEncrypt(sSQL)
        'Set Records = OpenSilentHttpCommsRS(sString, True)
        Set Records = OpenServerRSCompressed(WebSite & "oasis4.asp", "id", sSQL)
    
112     If Not Records.Fields.Item("SettingTablePrefix").value = "" Then
114         'sString = WebSite & "oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM " & Records.Fields.Item("SettingTablePrefix").Value & "AppSettings")
            'Set RSAppSetting = OpenSilentHttpCommsRS(sString, True)
            Set RSAppSetting = OpenServerRSCompressed(WebSite & "oasis4.asp", "id", "SELECT * FROM " & Records.Fields.Item("SettingTablePrefix").value & "AppSettings")
        End If
    
120     RSAppSetting.Close
122     Records.Close
    
124     Set RSAppSetting = Nothing
126     Set Records = Nothing
    
        '<EhFooter>
        Exit Sub

ReNewAppSettings_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.ReNewAppSettings " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub UpdateAppSettings(sVals As String)
'        '<EhHeader>
'        On Error GoTo UpdateAppSettings_Err
'        '</EhHeader>
'        Dim AllSettings() As String
'        Dim Setting() As String
'        Dim i As Integer
'        Dim sID As String
'        Dim sMySubQuery As String
'        Dim sString As String
'
'100     AllSettings = Split(sVals, vbCrLf)
'
'102     For i = LBound(AllSettings) To UBound(AllSettings) - 1
'104         Setting = Split(AllSettings(i), " = ")
'
'106         If i = 0 Then
'108             sID = Setting(1)
'110         ElseIf i = 1 Then
'112             sID = Setting(1)
'            Else
'                sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt(Replace(sID, "'", "") & "&FieldName=" & Replace(Setting(0), "'", ""))
'118             sMySubQuery = OpenSilentHttpCommsResponse(sString, True)
'
'            End If
'
'        Next
'
'        '<EhFooter>
'        Exit Sub
'
'UpdateAppSettings_Err:
'        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.UpdateAppSettings " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SaveStartUpParams
    On Error Resume Next
    m_HK.UnregisterKey "Killer"
    m_HK.UnregisterKey "LoginD"
    m_HK.UnregisterKey "Sync"
    m_HK.UnregisterKey "NewDB"
    m_HK.UnregisterKey "ConnSpeed"
    m_HK.UnregisterKey "Debug"
    
    Set m_HK = Nothing
    SystrayOff Me
    Me.Visible = False
        
    If OK Then
        OK = False
        LoadMainWin
    Else
        Unload frmSplash
        Unload m_frmDebug
        Set m_frmDebug = Nothing
    End If

End Sub

Private Sub SaveStartSettings(Optional bNewStart As Boolean)
        '<EhHeader>
        On Error GoTo SaveStartSettings_Err
        '</EhHeader>
        Dim FileName1 As String
        Dim Txt As String
        Dim i As Integer
        Dim Phrase As String, Position As Integer, Asc1 As Long, Char1 As String
        Dim lPhraseElement As Long

100     FileName1 = g_sAppPath & "\data\user\Sessions\start.dat"
            
102     If bNewStart Then
        
            Dim sServerURL As String
'tryagain:
104         'sServerURL = InputBox("It seems like this is the first time you are using this OASIS Client." & vbCrLf & "You need to enter the OASIS server address provided by your OASIS administrator.", "OASIS Client Settings", "atlantis.oasiswebservice.org")
            
            MsgBox "It seems like this is the first time you are using this OASIS Client." & vbCrLf & "Please select your respective OASIS Cloud Server." & vbCrLf & vbCrLf & "The atlantis server has been selected as default.", vbInformation, "OASIS Client Settings"
            sServerURL = "atlantis.oasiswebservice.org"
            g_sAppServerPath = sServerURL
            Call cmdExpand_Click
            c1Tab.CurrTab = 2
            listServer.Text = sServerURL
            
106         'If sServerURL = "" Then
108          '   MsgBox "The OASIS Server Address was wrong. Please try Again.", vbInformation, "OASIS Client Login Failure"
110           '  GoTo tryagain
            'End If
            
            ComServer.Text = sServerURL
            g_sAppServerPath = "http://" & ComServer.Text
            cmdRestoreSQL.Enabled = IIf(FileExists(g_sAppPath & "\data\db\" & Replace(g_sAppServerPath, "/", "-") & ".bak"), True, False)
            
112         'If MsgBox("If you do not have an OASIS account or internet connection press Cancel." & vbCrLf & "You will still be able to use OASIS but only in disconnected mode.", vbOKCancel, "OASIS Installation Routines") <> vbOK Then
114          '   g_bDemoLogin = True
            'End If

            Phrase = "prev=2" & vbCrLf
116         'Phrase = IIf(g_bDemoLogin, "prev=2", "prev=0") & vbCrLf
118         Phrase = Phrase & "url1=http://" & sServerURL & "/oasis4.asp" & vbCrLf

           ' GetClientDatabasePassword "http://" & sServerURL

           ' If g_ClientDBPassword = "none" Then

120             Phrase = Phrase & "user = john" & vbCrLf 'user = admin
122             Phrase = Phrase & "pass = doe" & vbCrLf 'pass = admin
          '  Else
            '    Phrase = Phrase & "user = password" & vbCrLf 'user = admin
            '    Phrase = Phrase & "pass = " & g_ClientDBPassword & vbCrLf 'pass = admin
'
          '  End If

124         Phrase = Phrase & "appserver=http://" & sServerURL & vbCrLf 'appserver=http://www.immap.org
126         Phrase = Phrase & "Provider=Microsoft.Jet" 'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data

            
            Dim oIni As New clIniReader
    
            CreateNewINI g_sAppPath & "\data\user\Sessions\sup.ini", oIni
    
            With oIni
                .Path = g_sAppPath & "\data\user\Sessions\sup.ini"
                .Section = "default"
            
                .Key = "Servers"
                .value = sServerURL
                .AddKeyWithValue

                .Key = "DefServer"
                .value = ComServer.Text
                .AddKeyWithValue
                
                .Key = "Database"
                .value = ComDatabase.Text
                .AddKeyWithValue
                
                .Key = "RememberMe"
                .value = IIf(chkRememberUser.value = vbChecked, "true", "false")
                .AddKeyWithValue
                
            End With

        Else
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'for some reasons the "Provider=" element was added each time you login
            'and causes an overflow after 52 logins - this cleans that up
            
            lPhraseElement = UBound(m_LoginParams)

            Do Until lPhraseElement = 1
            
                If m_LoginParams(lPhraseElement) = m_LoginParams(lPhraseElement - 1) Then
                    ReDim Preserve m_LoginParams(UBound(m_LoginParams) - 1)
                End If

                lPhraseElement = lPhraseElement - 1
            Loop

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
128         Phrase = IIf(bNewStart, "prev=0", "prev=1") & vbCrLf
    
130         m_LoginParams(UBound(m_LoginParams)) = m_Cnn.ConnectionString
                          
132         For i = LBound(m_LoginParams) + 1 To UBound(m_LoginParams)
134             Phrase = Phrase & m_LoginParams(i) & vbCrLf
            Next
        
        End If
    
136     For Position = Len(Phrase) To 1 Step -1
138         Char1 = Mid$(Phrase, Position, 1)
        
140         Asc1 = Asc(Char1)
        
142         Asc1 = (Asc1 * Asc1) / (Asc1 / 2)
        
144         Char1 = Chr$(Asc1)
                
146         Txt = Txt & Char1
        Next
    
148     Open FileName1 For Output As #1
150     Print #1, Txt
152     Close #1
    
154     If Not bNewStart Then m_bPrevLogin = True
        
156     If g_bDemoLogin Then m_bPrevLogin = True
    
158     g_sAppSettingsTable = "AppSettings"

        On Error Resume Next
        
        Dim ofs As New FileSystemObject
        
        If ofs.FileExists(g_sAppPath & "\OASIS_SynchNG_Client.exe") Then
            Shell g_sAppPath & "\OASIS_SynchNG_Client.exe /PermissionManagerInstall"
        End If

        Set ofs = Nothing

        '<EhFooter>
        Exit Sub

SaveStartSettings_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.SaveStartSettings " & "at line " & Erl
        'Resume Next
        '</EhFooter>
End Sub

Private Sub WipeTablesForReset(oConn As ADODB.Connection)
        '<EhHeader>
        On Error GoTo WipeTablesForReset_Err
        '</EhHeader>
        Dim RSUpdater As ADODB.Recordset
        
        On Error Resume Next
        
        Dim bDeleteIncidents As Boolean
        Dim oRS As New ADODB.Recordset
        
        bDeleteIncidents = True
        oRS.Open "SELECT top 1 * from [oincidents_FEA]", oConn, adOpenDynamic, adLockBatchOptimistic
bDeleteIncidents = True
        If oRS.State = adStateOpen Then
        
            If Not oRS.EOF And Not oRS.Bof Then
          
               ' bDeleteIncidents = IIf(MsgBox("Do you want to delete all incident data?", vbYesNo, "Confirm incident data deletion") = vbYes, True, False)
            'bDeleteIncidents = True
            End If
            
            oRS.Close
        
        End If
        
        Set oRS = Nothing
        
100     Set RSUpdater = New ADODB.Recordset

102     With RSUpdater
            
104         .Open "SELECT * FROM AppSettings", oConn, adOpenDynamic, adLockBatchOptimistic
106         .Find "SettingName = 'ProfileSettings'"

108         If Not .EOF Then
110             .Fields("SettingValue1").value = "-1"
                '.Fields("SettingValue2").Value = "0"
112             .Fields("SettingValue3").value = "-1"
114             .Fields("SettingValue4").value = "-1"
116             .Fields("SettingValue5").value = "-1"
118             .Fields("SettingValue6").value = "-1"
120             .Fields("SettingValue7").value = "-1"
122             .Fields("SettingValue8").value = "-1"
124             .Fields("SettingValue9").value = "-1"
126             .Fields("SettingValue10").value = "-1"
128             .UpdateBatch adAffectCurrent
130             '.Close
            End If
            
            If Not .Bof Or Not .EOF Then .MoveFirst
            .Find "SettingName = 'MapProjectDef'"

            If Not .EOF Then
                .Fields("SettingValue1").value = "-1"
                .Fields("SettingValue2").value = "-1"
                .Fields("SettingValue3").value = "-1"
                .Fields("SettingValue4").value = "-1"
                .Fields("SettingValue5").value = "-1"
                .Fields("SettingValue6").value = "-1"
                .Fields("SettingValue7").value = "-1"
                .Fields("SettingValue8").value = "-1"
                .Fields("SettingValue9").value = "-1"
                .Fields("SettingValue10").value = "-1"
                .UpdateBatch adAffectCurrent
                .Close
            End If
            
            

        End With

132     Set RSUpdater = Nothing
        oConn.Execute "delete from Attachments"

134     If bDeleteIncidents Then oConn.Execute "delete from oincidents_FEA"
136     If bDeleteIncidents Then oConn.Execute "delete from oincidents_GEO"
        oConn.Execute "delete from Incidents_ChartSettings"
        oConn.Execute "delete from ttkGISLayerSQLInProject"
        oConn.Execute "delete from ttkGISProjectDef WHERE bUGMap = true"
'140     oConn.Execute "delete from Charting"
'142     oConn.Execute "delete from ClientDBUpdates"
'144     oConn.Execute "delete from ClientFileLocations"
'146     oConn.Execute "delete from DataPacks"
148     oConn.Execute "delete from DynamicDataDefs"
150     oConn.Execute "delete from FeedGroups"
152     oConn.Execute "delete from Feeds"
154     oConn.Execute "delete from GeoBookMarks"
156     oConn.Execute "delete from GeoBookMarksCategories"
158     oConn.Execute "delete from GISGridTableSettings"
160     oConn.Execute "delete from Lang"
162     oConn.Execute "delete from Maps"
164     oConn.Execute "delete from PrintTemplates"
 oConn.Execute "delete from WebTiles"
 
'166     oConn.Execute "delete from SynchFeed"
'168     oConn.Execute "delete from SynchFeedsHistory"
'170     oConn.Execute "delete from SynchFiles"
'172     oConn.Execute "delete from SynchFolders"

174     If bDeleteIncidents Then
            oConn.Execute "delete from SynchHistory"
            oConn.Execute "delete from SynchHistoryOverview"
        Else
            oConn.Execute "delete from SynchHistory where [stablename] <> 'oincidents'"
            oConn.Execute "delete from SynchHistoryOverview where [stablename] <> 'oincidents'"
        End If
        
'176     oConn.Execute "delete from SynchTables"
        
178     Set RSUpdater = New ADODB.Recordset

180     With RSUpdater
            
182         .Open "SELECT top 1 * FROM [Personnell]", oConn, adOpenDynamic, adLockBatchOptimistic

186         If Not .EOF Then
188             .Fields("UserName").value = "bart"
190             .Fields("pwd").value = "simpson"
                '.Fields("Personnell_ID").Value = 2
192             .UpdateBatch adAffectCurrent
194             .Close
            End If

        End With

196     Set RSUpdater = Nothing
   
        '<EhFooter>
        Exit Sub

WipeTablesForReset_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.WipeTablesForReset " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub listServer_Click()

    If Len(listServer.Text) > 2 And Not listServer.Text = ComServer.Text Then
        cmdConnect.Enabled = True
    Else
        cmdConnect.Enabled = False
    End If

End Sub

Public Sub m_frmForcedLogin_DoForcedLogin(sUserName As String, _
                                          sPassword As String)
        '<EhHeader>
        On Error GoTo m_frmForcedLogin_DoForcedLogin_Err
        '</EhHeader>
                                          
        Dim sString As String
        Dim sResult As String
        
100     m_frmForcedLogin.cmdOK.Enabled = False
102     m_frmForcedLogin.cmdCancel.Enabled = False
    
104     InitialiseConnection
                                          
106     'sString = g_sAppServerPath & "\oasis4.asp" & "?user=" & CheckEncrypt(sUserName) & "&" & "pwd=" & CheckEncrypt(sPassword)
108     'sResult = OpenSilentHttpCommsResponse(sString, True)
        sResult = OpenServerResponseCompressed(g_sAppServerPath & "\oasis4.asp", "user", sUserName & "|||" & sPassword)

110     If Len(sResult) > 400 Then

112         GetRemoteTablePrefix sUserName, sPassword, g_sAppServerPath & "\oasis4.asp"
114         m_frmForcedLogin.LoginSucceeded = True

        End If
        
116     m_frmForcedLogin.cmdOK.Enabled = True
118     m_frmForcedLogin.cmdCancel.Enabled = True

        '<EhFooter>
        Exit Sub

m_frmForcedLogin_DoForcedLogin_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLogin.m_frmForcedLogin_DoForcedLogin " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmUpdateSettings_DoSynch(Index As Integer)
    
    Dim CN As ADODB.Connection
    Dim ORSLocalAppSettings As ADODB.Recordset
    Dim rsRemote As ADODB.Recordset
    Dim RS As ADODB.Recordset
    
    Set m_frmForcedLogin = New frmForcedLogin
    m_frmForcedLogin.Show vbModal, Me
        
    If Not m_frmForcedLogin.LoginSucceeded Then
    
        MsgBox "Login is needed to be able to do Synch/Updates", vbInformation
        
    Else

        If Index > 0 Then
    
            Set CN = New ADODB.Connection
            'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\OasisClient.mdb" & ";"
            CN.Open GetConnectionString(g_sAppPath & "\data\db\OasisClient.mdb")

            If Index > 1 Then
                Set rsRemote = New ADODB.Recordset
            End If
        
        End If

        Select Case Index
    
            Case 0
                ' Full Forced Synch
            
                Set ORSLocalAppSettings = New ADODB.Recordset
                ORSLocalAppSettings.Open "SELECT * FROM AppSettings", m_Cnn, adOpenDynamic, adLockReadOnly
        
                CheckProfileUpdate g_sUserName, g_sUserPass, g_sAppServerPath & "/oasis4.asp", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, RS

            Case 1
                'Application Settings
                Set ORSLocalAppSettings = New ADODB.Recordset
                ORSLocalAppSettings.Open "SELECT * FROM AppSettings", m_Cnn, adOpenDynamic, adLockReadOnly
            
                If g_sRemoteTablePrefix = "" Then
                    GetRemoteTablePrefix g_sUserName, g_sUserPass, g_sAppServerPath & "/oasis4.asp"
                End If
        
                GetRemoteProfile g_sAppServerPath & "/oasis4.asp", 0, g_sUserName, g_sUserPass
                AppSettingsUpdate CN, ORSLocalAppSettings
                MsgBox "Application Settings have been updated"

            Case 2
                'GIS Attribute grid
                UGTableUpdate "GISGridTableSettings", CN, "SettingValue4", "ProfileSettings"
                'GISGridUpdate CN, RSRemote, RS
                MsgBox "GIS Attribute Grid Settings have been updated"
                
            Case 3
                'Synchronisation Layer
                'OLD CODE
                'SynchLayerUpdate CN, RSRemote, RS, 0
                'MsgBox "Synch Layer Settings have been updated"

            Case 4
                'GeoMarks are updated by the SynchRunner
                    
            Case 5
                ' Print Templates
            
            Case 6

                ' Map Products
            Case 7
                'Auto Update
                '                ShellExecute Me.Hwnd, vbNullString, g_sAppPath & "\OASIS_SynchNG_Client.exe", "CheckBackground", "C:\", 1
                '                ShellExecute Me.Hwnd, vbNullString, g_sAppPath & "\AUClient.exe", "CheckBackground", "C:\", 1

            Case 8
                ' Charts
            Case 9
                UGTableUpdate "Themes", CN, "SettingValue10", "ProfileSettings"
                UGTableUpdate "ThemeGroups", CN, "", ""
                'ThematicsUpdate CN, RSRemote, RS
                MsgBox "Thematic Settings have been updated"

            Case 10
                ' Feeds / Dynamic Data
            Case 11
                'DynamDataDefsUpdate CN, RSRemote, RS
                UGTableUpdate "DynamicDataDefs", CN, "SettingValue7", "ProfileSettings"
                MsgBox "Dynamic Data Settings have been updated"
        End Select

    End If

    On Error Resume Next
    CN.Close
    Set CN = Nothing
    Set RS = Nothing
    Set rsRemote = Nothing
    Set ORSLocalAppSettings = Nothing
    
End Sub

Private Sub KillerOnTheLoose(Optional OnlyResetTbl As Boolean)
    On Error Resume Next
    Dim ofs As New FileSystemObject
    Dim bRestoreMSSQL As Boolean
    Dim bSQLServerExists As Boolean

    If Not OnlyResetTbl Then
        Kill g_sAppPath & "\data\user\Sessions\start.dat"
        Kill g_sAppPath & "\data\user\Sessions\main.dat"
        Kill g_sAppPath & "\data\user\Sessions\guid.dat"
    End If
    
    Dim CN As ADODB.Connection
    bRestoreMSSQL = False
    bSQLServerExists = MSSQL_CheckIfInstalled
     
    If bSQLServerInUse And FileExists(g_sAppPath & "\data\db\" & Replace(frmLogin.ComServer.Text, "/", "-") & ".bak") Then
        bRestoreMSSQL = IIf(MsgBox("Do you want to restore from a backup?", vbYesNo) = vbYes, True, False)
    End If
           
    If bSQLServerInUse And bRestoreMSSQL Then
        MSSQL_RestoreFromBackup
    ElseIf Not bSQLServerInUse Then
        Set CN = New ADODB.Connection
        CN.Open GetConnectionString(g_sAppPath & "\data\db\OasisClient.mdb")
        WipeTablesForReset CN
        CN.Close
        Set CN = Nothing
    End If
    
    Set ofs = Nothing
End Sub

Private Sub m_HK_HotKeyPress(ByVal sName As String, _
                             ByVal eModifiers As EHKModifiers, _
                             ByVal eKey As KeyCodeConstants)
Dim CN As ADODB.Connection

    If sName = "Killer" Then
        
        FormOnTopEx frmLogin.hwnd, False

        If MsgBox("You are about to Reset the OASIS Client. Are you sure?", vbYesNo, "OASIS Client") = vbYes Then
            On Error Resume Next
            KillerOnTheLoose
            OK = False
            Unload Me 'Me.Hide
        Else

            FormOnTop Me
        End If
    
    ElseIf sName = "NewDB" Then
        'mnuOpenDB_Click
    ElseIf sName = "LoginD" Then
        MsgBox "Command Line: " & Command$ & vbCrLf & "Current Used Path:" & g_sAppPath & "\data\db\" & vbCrLf & "Revision Version: " & App.Revision & " " & App.Comments, vbInformation, "OASIS VALUES"
    ElseIf sName = "Sync" Then
    
        Dim lWidth As Long
        Dim lHeight As Long
        Dim lLeft As Long
    
        lWidth = Me.Width
        lHeight = Me.Height
        lLeft = Me.left
    
        Me.left = -4000
        Me.Height = 2
        Me.Width = 2
        
        Set m_frmUpdateSettings = New frmUpdateSettings
        m_frmUpdateSettings.Show vbModal, Me
        
        Me.Height = lHeight
        Me.Width = lWidth
        Me.left = lLeft

        Unload m_frmUpdateSettings
        
    ElseIf sName = "ConnSpeed" Then
    
        Set CN = New ADODB.Connection
        'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\OasisClient.mdb" & ";"
        CN.Open GetConnectionString(g_sAppPath & "\data\db\OasisClient.mdb")
        FormOnTopEx Me.hwnd, False
        SetConnectionSpeed CN

        FormOnTop Me
                 
        CN.Close
        Set CN = Nothing
        
    ElseIf sName = "Debug" Then
    
        m_frmDebug.Show
    End If

End Sub

Private Sub SetConnectionSpeed(oCnn As ADODB.Connection)
        '<EhHeader>
        On Error GoTo SetConnectionSpeed_Err
        '</EhHeader>

        Dim m_frmDialogWithTwoFields As frmDialogWithTwoFields
100     Set m_frmDialogWithTwoFields = New frmDialogWithTwoFields

        Dim iTimeout As Integer
        Dim iRetries As Integer

        Dim oRS As ADODB.Recordset
102     Set oRS = New ADODB.Recordset
104     oRS.Open "SELECT SettingName, SettingValue1, SettingValue2 FROM AppSettings WHERE SettingName = 'ServerConnectionParameters'", oCnn, adOpenDynamic, adLockBatchOptimistic

106     m_frmDialogWithTwoFields.caption = "Server Connection Parameters"
108     m_frmDialogWithTwoFields.lbl1.caption = "Server connection timeout (in seconds):"
110     m_frmDialogWithTwoFields.lbl2.caption = "Server connection retries"

112     If oRS.EOF Then
114         iTimeout = 20
116         iRetries = 3
118         m_frmDialogWithTwoFields.txt1.Text = iTimeout
120         m_frmDialogWithTwoFields.txt2.Text = iRetries
122         m_frmDialogWithTwoFields.Show vbModal, Me
124         iTimeout = IIf(IsNumeric(m_frmDialogWithTwoFields.sText1), m_frmDialogWithTwoFields.sText1, iTimeout)
126         iRetries = IIf(IsNumeric(m_frmDialogWithTwoFields.sText2), m_frmDialogWithTwoFields.sText2, iRetries)
128         oRS.AddNew
130         oRS.Fields(0).value = "ServerConnectionParameters"
        Else
132         iTimeout = IIf(IsNumeric(oRS.Fields(1).value), oRS.Fields(1).value, 20)
134         iRetries = IIf(IsNumeric(oRS.Fields(2).value), oRS.Fields(2).value, 3)
136         m_frmDialogWithTwoFields.txt1.Text = iTimeout
138         m_frmDialogWithTwoFields.txt2.Text = iRetries
140         m_frmDialogWithTwoFields.Show vbModal, Me
142         iTimeout = IIf(IsNumeric(m_frmDialogWithTwoFields.sText1), m_frmDialogWithTwoFields.sText1, iTimeout)
144         iRetries = IIf(IsNumeric(m_frmDialogWithTwoFields.sText2), m_frmDialogWithTwoFields.sText2, iRetries)
        End If
    
146     If m_frmDialogWithTwoFields.bClickedOK Then
148         oRS.Fields(1).value = iTimeout
150         oRS.Fields(2).value = iRetries
152         oRS.UpdateBatch adAffectCurrent
          '  m_frmOASISProgress.SetClientDBTimeoutSettings CStr(iTimeout), CStr(iRetries), g_sAppPath & "\data\db\OasisClient.mdb"
           ' SetServerCommsSettings CStr(iTimeout), CStr(iRetries), g_sAppPath & "\data\db\OasisClient.mdb"
        End If

154     oRS.Close
156     Set oRS = Nothing
158     Set m_frmDialogWithTwoFields = Nothing

        '<EhFooter>
        Exit Sub

SetConnectionSpeed_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLogin.SetConnectionSpeed " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub



Sub MainCommand()
        '<EhHeader>
        On Error GoTo MainCommand_Err
        '</EhHeader>
        Dim a_strArgs() As String
        Dim blnDebug As Boolean
        Dim strFileName As String
        Dim mFileSysObj As FileSystemObject
        Dim i As Integer
   
100     a_strArgs = Split(Command$, " ")

102     For i = LBound(a_strArgs) To UBound(a_strArgs)

104         Select Case LCase(a_strArgs(i))

                Case "-f"
                    mFileSysObj = New FileSystemObject
106                 mFileSysObj.DeleteFile g_sAppPath & "\data\user\Sessions\start.dat", True
108                 DebugPrint a_strArgs(i + 1)
110                 i = i + 1

112             Case "-u"
114                 txtUserName.Text = a_strArgs(i + 1)
116                 i = i + 1

118             Case "-p"
120                 txtPassword.Text = a_strArgs(i + 1)
122                 i = i + 1

124             Case Else
                    'MsgBox "Invalid argument: " & a_strArgs(i)
            End Select
      
126     Next i
        
        '<EhFooter>
        Exit Sub

MainCommand_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.MainCommand " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub OptMicrosoftAccess_Click()

    If OptMicrosoftAccess.value = True Then
        ComDatabase.Text = OptMicrosoftAccess.caption
        ComServer.Text = sInitialMSACCESSServer
        listServer.Text = sInitialMSACCESSServer
        cmdConnect.Enabled = False
    End If
    
    bSQLServerInUse = False
    
End Sub

Private Sub OptMicrosoftSQL_Click()
'If OptMicrosoftAccess.Value = True Then ComDatabase.Text = OptMicrosoftAccess.caption
    If OptMicrosoftSQL.value = True Then ComDatabase.Text = OptMicrosoftSQL.caption
    bSQLServerInUse = True
End Sub




