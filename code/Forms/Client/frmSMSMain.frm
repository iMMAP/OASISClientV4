VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmSMSMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS SMS Messenger"
   ClientHeight    =   7425
   ClientLeft      =   1305
   ClientTop       =   1635
   ClientWidth     =   11820
   Icon            =   "frmSMSMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11820
   Begin VB.Data dtaTemplates 
      Caption         =   "dtaTemplates"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data dtaNetworks 
      Caption         =   "dtaNetworks"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data dtaCountrys 
      Caption         =   "dtaCountrys"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data dtaSettings 
      Caption         =   "dtaSettings"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1740
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6780
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data dtaPhonebook 
      Caption         =   "dtaPhoneBook"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog cmdDlgOpen 
      Left            =   7440
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin TabDlg.SSTab TabMain 
      Height          =   7335
      Left            =   60
      TabIndex        =   157
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   12
      TabsPerRow      =   12
      TabHeight       =   882
      TabCaption(0)   =   " SMS Dashboard"
      TabPicture(0)   =   "frmSMSMain.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblOriginator"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPassword"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblUserkey"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCurrentBalance"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraCurrentJob"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraOptions"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraSMSType"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtOriginator"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPassword"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtUserkey"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "FraActiveMessage"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "pctLogo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Recipients"
      TabPicture(1)   =   "frmSMSMain.frx":D0B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPhonebook"
      Tab(1).Control(1)=   "fraCurrentRecipients"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Text SMS"
      TabPicture(2)   =   "frmSMSMain.frx":13916
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraTemplates"
      Tab(2).Control(1)=   "fraCurrentMessage"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Operator-/ Grouplogos"
      TabPicture(3)   =   "frmSMSMain.frx":1A178
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdSelectHandylogo"
      Tab(3).Control(1)=   "txtPathLogo"
      Tab(3).Control(2)=   "cmdEditLogo"
      Tab(3).Control(3)=   "fraNetworkSettings"
      Tab(3).Control(4)=   "fraRandomLogos"
      Tab(3).Control(5)=   "Inet1"
      Tab(3).Control(6)=   "winsckRandomLogo(0)"
      Tab(3).Control(7)=   "winsckRandomLogo(1)"
      Tab(3).Control(8)=   "winsckRandomLogo(2)"
      Tab(3).Control(9)=   "winsckRandomLogo(3)"
      Tab(3).Control(10)=   "winsckRandomLogo(4)"
      Tab(3).Control(11)=   "winsckRandomLogo(5)"
      Tab(3).Control(12)=   "winsckRandomLogo(6)"
      Tab(3).Control(13)=   "winsckRandomLogo(7)"
      Tab(3).Control(14)=   "winsckRandomLogo(8)"
      Tab(3).Control(15)=   "winsckRandomLogo(9)"
      Tab(3).Control(16)=   "winsckRandomLogo(10)"
      Tab(3).Control(17)=   "winsckRandomLogo(11)"
      Tab(3).Control(18)=   "winsckRandomLogo(12)"
      Tab(3).Control(19)=   "winsckRandomLogo(13)"
      Tab(3).Control(20)=   "winsckRandomLogo(14)"
      Tab(3).Control(21)=   "winsckRandomLogo(15)"
      Tab(3).Control(22)=   "winsckRandomLogo(16)"
      Tab(3).Control(23)=   "winsckRandomLogo(17)"
      Tab(3).Control(24)=   "winsckRandomLogo(18)"
      Tab(3).Control(25)=   "winsckRandomLogo(19)"
      Tab(3).Control(26)=   "winsckRandomLogo(20)"
      Tab(3).Control(27)=   "winsckLogoControl"
      Tab(3).Control(28)=   "imgPreviewOperatorLogo"
      Tab(3).Control(29)=   "lblPathLogo"
      Tab(3).ControlCount=   30
      TabCaption(4)   =   "Ringtones"
      TabPicture(4)   =   "frmSMSMain.frx":1A194
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdSelectRingtone"
      Tab(4).Control(1)=   "txtPathRingtone"
      Tab(4).Control(2)=   "lblPathRingtone"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Picturemessage"
      TabPicture(5)   =   "frmSMSMain.frx":1A1B0
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdSelectPictureMessage"
      Tab(5).Control(1)=   "txtPathPictureMessage"
      Tab(5).Control(2)=   "txtPictureMessageText"
      Tab(5).Control(3)=   "cmdEditPictureMessage"
      Tab(5).Control(4)=   "lblPathPicture"
      Tab(5).Control(5)=   "lblCounterPictureMessage"
      Tab(5).Control(6)=   "imgPreviewPictureMessage"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "VCard"
      TabPicture(6)   =   "frmSMSMain.frx":1A1CC
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "chkVCardBlinkingSMS"
      Tab(6).Control(1)=   "txtVCardPhoneNumber"
      Tab(6).Control(2)=   "txtVCardName"
      Tab(6).Control(3)=   "lblVCardPhoneNumber"
      Tab(6).Control(4)=   "lblVCardName"
      Tab(6).ControlCount=   5
      TabCaption(7)   =   "Binarydata"
      TabPicture(7)   =   "frmSMSMain.frx":1A1E8
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "txtXSer"
      Tab(7).Control(1)=   "txtMessageData"
      Tab(7).Control(2)=   "lblXSer"
      Tab(7).Control(3)=   "lblMessagedata"
      Tab(7).ControlCount=   4
      TabCaption(8)   =   "Waiting Indication"
      TabPicture(8)   =   "frmSMSMain.frx":1A204
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lblCountVoiceMessages"
      Tab(8).Control(1)=   "lblCountFaxMessages"
      Tab(8).Control(2)=   "lblCountEMailMessages"
      Tab(8).Control(3)=   "lblCountOtherMessages"
      Tab(8).Control(4)=   "chkStoreMessage"
      Tab(8).Control(5)=   "txtMessageWaitingIndication"
      Tab(8).Control(6)=   "txtCountVoiceMessages"
      Tab(8).Control(7)=   "txtCountFaxMessages"
      Tab(8).Control(8)=   "txtCountEMailMessages"
      Tab(8).Control(9)=   "txtCountOtherMessages"
      Tab(8).Control(10)=   "chkVoiceMessages"
      Tab(8).Control(11)=   "chkFaxMessages"
      Tab(8).Control(12)=   "chkEmailMessages"
      Tab(8).Control(13)=   "chkOtherMessages"
      Tab(8).Control(14)=   "chkMessageWaitingIndicationBlinkingSMS"
      Tab(8).ControlCount=   15
      TabCaption(9)   =   "Unicode Text SMS"
      TabPicture(9)   =   "frmSMSMain.frx":1A220
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "fraCurrentMessageUnicode"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "WAP Push"
      TabPicture(10)  =   "frmSMSMain.frx":1A23C
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lblWAPPushSMSPictureDetails"
      Tab(10).Control(1)=   "lblWAPPushPictureDescription"
      Tab(10).Control(2)=   "lblPathWAPPushPicture"
      Tab(10).Control(3)=   "txtWAPPushSMSDescription"
      Tab(10).Control(4)=   "cmdSelectWAPPushSMSPicture"
      Tab(10).Control(5)=   "txtWAPPushSMSPicturePath"
      Tab(10).Control(6)=   "cmdEditWAPPushSMSPicture"
      Tab(10).Control(7)=   "picContainer"
      Tab(10).Control(8)=   "cmdHelpWAPPush"
      Tab(10).Control(9)=   "cmdTipsAndTricksWAPPush"
      Tab(10).ControlCount=   10
      TabCaption(11)  =   "Work Logging"
      TabPicture(11)  =   "frmSMSMain.frx":1A258
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "fraJobRemarks"
      Tab(11).Control(1)=   "fraAdditionalInformation"
      Tab(11).Control(2)=   "fraDeliveryNotificationsPerSMS"
      Tab(11).ControlCount=   3
      Begin VB.PictureBox pctLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   5880
         ScaleHeight     =   405
         ScaleWidth      =   5565
         TabIndex        =   169
         Top             =   6780
         Width           =   5595
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   60
            Picture         =   "frmSMSMain.frx":20ABA
            ScaleHeight     =   420
            ScaleWidth      =   5355
            TabIndex        =   170
            Top             =   0
            Width           =   5355
         End
      End
      Begin VB.Frame FraActiveMessage 
         Caption         =   "Active message:"
         Height          =   1455
         Left            =   120
         TabIndex        =   164
         Top             =   4560
         Width           =   5655
         Begin VB.TextBox txtSMSPreview 
            Enabled         =   0   'False
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   165
            Top             =   240
            Width           =   5415
         End
      End
      Begin VB.CommandButton cmdTipsAndTricksWAPPush 
         Caption         =   "Tips and Tricks..."
         Height          =   375
         Left            =   -65040
         TabIndex        =   123
         Top             =   920
         Width           =   1575
      End
      Begin VB.CommandButton cmdHelpWAPPush 
         Caption         =   "Help..."
         Height          =   375
         Left            =   -66720
         TabIndex        =   122
         Top             =   920
         Width           =   1575
      End
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   -74880
         ScaleHeight     =   4065
         ScaleWidth      =   5505
         TabIndex        =   124
         Top             =   1560
         Width           =   5535
         Begin VB.PictureBox picHideBottomRightCorner 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   4920
            ScaleHeight     =   345
            ScaleWidth      =   345
            TabIndex        =   162
            Top             =   3360
            Width           =   375
         End
         Begin VB.VScrollBar scrlVertical 
            Height          =   1215
            Left            =   5160
            TabIndex        =   161
            TabStop         =   0   'False
            Top             =   120
            Width           =   255
         End
         Begin VB.HScrollBar scrlHorizontal 
            Height          =   255
            Left            =   240
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Image imgPreviewWAPPushSMSPicture 
            Appearance      =   0  'Flat
            Height          =   135
            Left            =   0
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.CommandButton cmdSelectPictureMessage 
         Caption         =   "..."
         Height          =   255
         Left            =   -66960
         TabIndex        =   84
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtPathPictureMessage 
         Height          =   285
         Left            =   -72600
         TabIndex        =   83
         Top             =   480
         Width           =   5535
      End
      Begin VB.TextBox txtPictureMessageText 
         Height          =   1935
         Left            =   -74880
         MaxLength       =   121
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   86
         Text            =   "frmSMSMain.frx":2803C
         Top             =   840
         Width           =   3615
      End
      Begin VB.CommandButton cmdSelectRingtone 
         Caption         =   "..."
         Height          =   255
         Left            =   -66960
         TabIndex        =   81
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtPathRingtone 
         Height          =   285
         Left            =   -72600
         TabIndex        =   80
         Top             =   480
         Width           =   5535
      End
      Begin VB.CommandButton cmdSelectHandylogo 
         Caption         =   "..."
         Height          =   255
         Left            =   -66960
         TabIndex        =   63
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtPathLogo 
         Height          =   285
         Left            =   -72600
         TabIndex        =   62
         Top             =   480
         Width           =   5535
      End
      Begin VB.TextBox txtUserkey 
         Height          =   285
         Left            =   9840
         MaxLength       =   12
         TabIndex        =   1
         Top             =   1020
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   8040
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtOriginator 
         Height          =   285
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   5
         Top             =   660
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame fraSMSType 
         Caption         =   "SMS Type"
         Height          =   3675
         Left            =   120
         TabIndex        =   7
         Top             =   900
         Width           =   5655
         Begin VB.OptionButton optSMSType 
            Caption         =   "Text SMS"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   5415
         End
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Controller"
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         Top             =   6000
         Width           =   5655
         Begin VB.CommandButton cmdGeneralSettings 
            Caption         =   "General Settings..."
            Height          =   855
            Left            =   120
            Picture         =   "frmSMSMain.frx":2808E
            Style           =   1  'Graphical
            TabIndex        =   163
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send!"
            Enabled         =   0   'False
            Height          =   855
            Left            =   3840
            Picture         =   "frmSMSMain.frx":2E8E0
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdDeferredDeliveryTime 
            Caption         =   "Adjust Delivery Time..."
            Height          =   855
            Left            =   1980
            Picture         =   "frmSMSMain.frx":35132
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkVCardBlinkingSMS 
         Caption         =   "Blinking SMS"
         Height          =   255
         Left            =   -74880
         TabIndex        =   91
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtVCardPhoneNumber 
         Height          =   285
         Left            =   -73080
         TabIndex        =   90
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtVCardName 
         Height          =   285
         Left            =   -73080
         TabIndex        =   88
         Top             =   480
         Width           =   2415
      End
      Begin VB.CheckBox chkMessageWaitingIndicationBlinkingSMS 
         Caption         =   "Blinking SMS"
         Height          =   255
         Left            =   -74880
         TabIndex        =   109
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CheckBox chkOtherMessages 
         Caption         =   "Other Messages"
         Height          =   255
         Left            =   -74880
         TabIndex        =   105
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkEmailMessages 
         Caption         =   "EMail Messages"
         Height          =   255
         Left            =   -74880
         TabIndex        =   102
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkFaxMessages 
         Caption         =   "Fax Messages"
         Height          =   255
         Left            =   -74880
         TabIndex        =   99
         Top             =   840
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkVoiceMessages 
         Caption         =   "Voice Messages"
         Height          =   255
         Left            =   -74880
         TabIndex        =   96
         Top             =   480
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.TextBox txtCountOtherMessages 
         Height          =   285
         Left            =   -69960
         TabIndex        =   107
         Text            =   "255"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtCountEMailMessages 
         Height          =   285
         Left            =   -69960
         TabIndex        =   104
         Text            =   "255"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtCountFaxMessages 
         Height          =   285
         Left            =   -69960
         TabIndex        =   101
         Text            =   "255"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtCountVoiceMessages 
         Height          =   285
         Left            =   -69960
         TabIndex        =   98
         Text            =   "255"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtMessageWaitingIndication 
         Height          =   1935
         Left            =   -68880
         MaxLength       =   121
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   110
         Text            =   "frmSMSMain.frx":3B984
         Top             =   480
         Width           =   3615
      End
      Begin VB.CheckBox chkStoreMessage 
         Caption         =   "Store Message in Mobilephone"
         Height          =   255
         Left            =   -74880
         TabIndex        =   108
         Top             =   2040
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.TextBox txtXSer 
         Height          =   285
         Left            =   -73320
         TabIndex        =   95
         Text            =   "010706050415821582"
         Top             =   1920
         Width           =   4815
      End
      Begin VB.TextBox txtMessageData 
         Height          =   1365
         Left            =   -73320
         MultiLine       =   -1  'True
         TabIndex        =   93
         Text            =   "frmSMSMain.frx":3B9D3
         Top             =   480
         Width           =   4815
      End
      Begin VB.Frame fraPhonebook 
         Caption         =   "Available Recipients"
         Height          =   6495
         Left            =   -69000
         TabIndex        =   40
         Top             =   600
         Width           =   5535
         Begin VB.TextBox txtPhonebookName 
            Height          =   285
            Left            =   1440
            MaxLength       =   150
            TabIndex        =   24
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtPhonebookPhoneNumber 
            Height          =   285
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   26
            Top             =   600
            Width           =   1935
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Save Edits"
            Height          =   795
            Left            =   3600
            Picture         =   "frmSMSMain.frx":3BAE2
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1020
            Width           =   1815
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
            Height          =   795
            Left            =   3600
            Picture         =   "frmSMSMain.frx":42334
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1860
            Width           =   1815
         End
         Begin VB.CommandButton cmdAddNew 
            Caption         =   "Add contact"
            Height          =   795
            Left            =   3600
            Picture         =   "frmSMSMain.frx":48B86
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   180
            Width           =   1815
         End
         Begin VB.CommandButton cmdAddRecipientsFromPhonebook 
            Caption         =   "<<"
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   2520
            Width           =   3255
         End
         Begin VB.TextBox txtFilter 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   34
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox txtPhonebookVariableField 
            Height          =   285
            Index           =   1
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   28
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtPhonebookVariableField 
            Height          =   285
            Index           =   2
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   30
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtPhonebookVariableField 
            Height          =   285
            Index           =   3
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   32
            Top             =   1680
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdPhonebook 
            Bindings        =   "frmSMSMain.frx":4F3D8
            Height          =   3255
            Left            =   120
            TabIndex        =   39
            Top             =   3120
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   5741
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            FixedCols       =   0
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            AllowUserResizing=   1
         End
         Begin VB.Label lblPhonebookname 
            Caption         =   "Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblPhonebookPhonenumber 
            Caption         =   "Phonenumber"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblFilter 
            Caption         =   "Search"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblPhonebookVariableField 
            Caption         =   "lblPhonebookVariableField"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblPhonebookVariableField 
            Caption         =   "lblPhonebookVariableField"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblPhonebookVariableField 
            Caption         =   "lblPhonebookVariableField"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   1680
            Width           =   1215
         End
      End
      Begin VB.Frame fraCurrentRecipients 
         Caption         =   "Current Recipients"
         Height          =   6495
         Left            =   -74880
         TabIndex        =   22
         Top             =   600
         Width           =   5775
         Begin VB.CommandButton cmdImportRecipientsFromFile 
            Caption         =   "Import Recipients from textfile"
            Height          =   615
            Left            =   3720
            TabIndex        =   18
            Top             =   5580
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddToRecipientList 
            Caption         =   "Add to worklist"
            Height          =   855
            Left            =   3720
            Picture         =   "frmSMSMain.frx":4F3F3
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtRecipient 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "Bitte internationales Nummernformat verwenden  z.B. +41791234567"
            Top             =   480
            Width           =   2775
         End
         Begin VB.CommandButton cmdClearList 
            Caption         =   "Clear List"
            Height          =   855
            Left            =   3720
            Picture         =   "frmSMSMain.frx":55C45
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   2040
            Width           =   1935
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove from worklist"
            Height          =   855
            Left            =   3720
            Picture         =   "frmSMSMain.frx":5C497
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1140
            Width           =   1935
         End
         Begin VB.CommandButton cmdDuplicates 
            Caption         =   "Search for duplicates"
            Height          =   855
            Left            =   3720
            Picture         =   "frmSMSMain.frx":62CE9
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   2940
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdRecipients 
            Height          =   5175
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   9128
            _Version        =   393216
            Rows            =   1
            Cols            =   5
            FixedCols       =   0
            ScrollTrack     =   -1  'True
            AllowUserResizing=   1
         End
         Begin VB.Label lblCurrentRecipients 
            Caption         =   "Current Recipients"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label lblEnterRecipientManually 
            Caption         =   "Empfänger manuell eingeben"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame fraCurrentJob 
         Caption         =   "Current Job"
         Height          =   5835
         Left            =   5880
         TabIndex        =   12
         Top             =   900
         Width           =   5655
         Begin VB.ListBox lstCurrentJob 
            Height          =   5325
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   5415
         End
      End
      Begin VB.CommandButton cmdEditLogo 
         Caption         =   "Edit"
         Height          =   255
         Left            =   -66360
         TabIndex        =   64
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditPictureMessage 
         Caption         =   "Edit"
         Height          =   255
         Left            =   -66360
         TabIndex        =   85
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame fraNetworkSettings 
         Caption         =   "Network Settings"
         Height          =   3015
         Left            =   -74880
         TabIndex        =   73
         Top             =   1320
         Width           =   6975
         Begin VB.TextBox txtMNC 
            Height          =   285
            Left            =   1320
            TabIndex        =   68
            Text            =   "1"
            Top             =   780
            Width           =   615
         End
         Begin VB.TextBox txtMCC 
            Height          =   285
            Left            =   1320
            TabIndex        =   66
            Text            =   "228"
            Top             =   360
            Width           =   615
         End
         Begin VB.ListBox lstCountrys 
            Height          =   1620
            Left            =   1320
            TabIndex        =   70
            Top             =   1200
            Width           =   2175
         End
         Begin VB.ListBox lstOperators 
            Height          =   1620
            Left            =   4680
            TabIndex        =   72
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label lblMNC 
            Caption         =   "MNC"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   780
            Width           =   855
         End
         Begin VB.Label lblMCC 
            Caption         =   "MCC"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblCountry 
            Caption         =   "Country"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblOperator 
            Caption         =   "Operator"
            Height          =   255
            Left            =   3600
            TabIndex        =   71
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.Frame fraRandomLogos 
         Caption         =   "Randomlogos / More than 650 logos available!"
         Height          =   4935
         Left            =   -67800
         TabIndex        =   78
         Top             =   1320
         Width           =   4095
         Begin VB.CommandButton cmdRandomLogoRefresh 
            Caption         =   "Load new randomlogos"
            Height          =   495
            Left            =   600
            TabIndex        =   75
            Top             =   3240
            Width           =   2895
         End
         Begin VB.CheckBox chkRandomLogo 
            Caption         =   "Check1"
            Height          =   255
            Left            =   480
            TabIndex        =   77
            Top             =   4560
            Width           =   3255
         End
         Begin VB.Frame fraSelectedRandomLogo 
            Caption         =   "Selected Randomlogo"
            Height          =   615
            Left            =   1080
            TabIndex        =   76
            Top             =   3840
            Width           =   2175
            Begin VB.Image imgSelectedRandomLogo 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   480
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   19
            Left            =   2760
            Top             =   720
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   17
            Left            =   2760
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   15
            Left            =   2760
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   13
            Left            =   2760
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   11
            Left            =   2760
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   9
            Left            =   1440
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   8
            Left            =   120
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   7
            Left            =   1440
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   6
            Left            =   120
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   5
            Left            =   1440
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   4
            Left            =   120
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   3
            Left            =   1440
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   120
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   1440
            Top             =   720
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   120
            Top             =   720
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   18
            Left            =   1440
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   16
            Left            =   2760
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   14
            Left            =   2760
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   12
            Left            =   1440
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   10
            Left            =   120
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Image imgRandomLogo 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   20
            Left            =   120
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label lblCopyright 
            Caption         =   "All items © 2000 - 2001 by Handylogos unlimited"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame fraTemplates 
         Caption         =   "Templates"
         Height          =   6495
         Left            =   -69360
         TabIndex        =   60
         Top             =   600
         Width           =   5895
         Begin VB.CommandButton cmdTemplateCopy 
            Caption         =   "<<"
            Height          =   495
            Left            =   120
            TabIndex        =   58
            Top             =   2040
            Width           =   495
         End
         Begin VB.CommandButton cmdTemplateAddnew 
            Caption         =   "Add to Templates"
            Height          =   735
            Left            =   3720
            Picture         =   "frmSMSMain.frx":6953B
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdTemplateDelete 
            Caption         =   "Delete"
            Height          =   735
            Left            =   3720
            Picture         =   "frmSMSMain.frx":6FD8D
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   1740
            Width           =   2055
         End
         Begin VB.CommandButton cmdTemplateSave 
            Caption         =   "Save"
            Height          =   735
            Left            =   3720
            Picture         =   "frmSMSMain.frx":765DF
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtTemplateMessage 
            Height          =   1725
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   52
            Top             =   240
            Width           =   3495
         End
         Begin MSFlexGridLib.MSFlexGrid grdTemplates 
            Bindings        =   "frmSMSMain.frx":7CE31
            Height          =   3615
            Left            =   180
            TabIndex        =   59
            Top             =   2760
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   6376
            _Version        =   393216
            Rows            =   1
            Cols            =   1
            FixedCols       =   0
            WordWrap        =   -1  'True
            ScrollTrack     =   -1  'True
         End
         Begin VB.Label lblNeededPartsTemplate 
            Caption         =   "lblNeededPartsTemplate"
            Height          =   255
            Left            =   720
            TabIndex        =   57
            Top             =   2280
            Width           =   4095
         End
         Begin VB.Label lblCharsLeftTemplate 
            Caption         =   "lblCharsLeftTemplate"
            Height          =   255
            Left            =   720
            TabIndex        =   56
            Top             =   2040
            Width           =   3195
         End
      End
      Begin VB.Frame fraCurrentMessage 
         Caption         =   "Current Message"
         Height          =   6495
         Left            =   -74880
         TabIndex        =   51
         Top             =   600
         Width           =   5415
         Begin VB.TextBox txtSMS 
            Height          =   1695
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Top             =   240
            Width           =   5175
         End
         Begin VB.CheckBox chkFlashingSMS 
            Caption         =   "Flash SMS"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   2520
            Width           =   1635
         End
         Begin VB.CheckBox chkBlinkingSMS 
            Caption         =   "Blinking SMS"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   2880
            Width           =   2655
         End
         Begin VB.CheckBox chkReplaceMessage 
            Caption         =   "Message may be replaced / will replace other messages"
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   3240
            Width           =   4695
         End
         Begin VB.CommandButton cmdInsertPlaceHolder 
            Caption         =   "Insert Placeholder"
            Height          =   795
            Left            =   1800
            Picture         =   "frmSMSMain.frx":7CE4C
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   2040
            Width           =   1695
         End
         Begin VB.ComboBox cboPlaceHolder 
            Height          =   315
            ItemData        =   "frmSMSMain.frx":8369E
            Left            =   120
            List            =   "frmSMSMain.frx":836B1
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   2070
            Width           =   1575
         End
         Begin VB.CommandButton cmdPreviewSMSJob 
            Caption         =   "Preview"
            Height          =   795
            Left            =   3600
            Picture         =   "frmSMSMain.frx":836F3
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label lblCharsLeftCurrentMessage 
            Caption         =   "lblCharsLeftCurrentMessage"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   3720
            Width           =   5175
         End
         Begin VB.Label lblNeededPartsCurrentMessage 
            Caption         =   "lblNeededPartsCurrentMessage"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   3960
            Width           =   5175
         End
         Begin VB.Label lblSMSPartsNote 
            Caption         =   "Please note: ......."
            Height          =   2055
            Left            =   120
            TabIndex        =   50
            Top             =   4320
            Visible         =   0   'False
            Width           =   5175
         End
      End
      Begin VB.Frame fraCurrentMessageUnicode 
         Caption         =   "Current Message"
         Height          =   6015
         Left            =   -74835
         TabIndex        =   115
         Top             =   375
         Width           =   4935
         Begin VB.TextBox txtUnicode 
            Height          =   2835
            Left            =   240
            TabIndex        =   168
            Top             =   300
            Width           =   4455
         End
         Begin VB.CheckBox chkReplaceMessageUnicode 
            Caption         =   "Message may be replaced / will replace other messages"
            Height          =   375
            Left            =   120
            TabIndex        =   113
            Top             =   3960
            Width           =   4695
         End
         Begin VB.CheckBox chkFlashingSMSUnicode 
            Caption         =   "Flash SMS"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   3600
            Width           =   2775
         End
         Begin VB.Label lblUCSNote 
            Caption         =   "Please note: ......."
            Height          =   1455
            Left            =   120
            TabIndex        =   114
            Top             =   4440
            Width           =   4575
         End
         Begin VB.Label lblCharsLeftCurrentMessageUnicode 
            Caption         =   "lblCharsLeftCurrentMessageUnicode"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   3240
            Width           =   4095
         End
      End
      Begin VB.Frame fraDeliveryNotificationsPerSMS 
         Caption         =   "Delivery Notifications per SMS"
         Height          =   2895
         Left            =   -69120
         TabIndex        =   147
         Top             =   600
         Width           =   5655
         Begin VB.CheckBox chkUseOTADeliveryNotifications 
            Caption         =   "Empfangsbestätigungen per SMS verwenden"
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   240
            Value           =   1  'Checked
            Width           =   4815
         End
         Begin VB.CheckBox chkDeliveryNotificationBuffered 
            Caption         =   "chkDeliveryNotificationBuffered"
            Height          =   615
            Left            =   120
            TabIndex        =   144
            Top             =   960
            Width           =   5175
         End
         Begin VB.CheckBox chkDeliveryNotificationDelivered 
            Caption         =   "chkDeliveryNotificationDelivered"
            Height          =   615
            Left            =   120
            TabIndex        =   145
            Top             =   1560
            Width           =   5295
         End
         Begin VB.CheckBox chkDeliveryNotificationNotDelivered 
            Caption         =   "chkDeliveryNotificationNotDelivered"
            Height          =   615
            Left            =   120
            TabIndex        =   146
            Top             =   2160
            Width           =   5295
         End
         Begin VB.TextBox txtRecipientDeliveryNotification 
            Height          =   285
            Left            =   2040
            TabIndex        =   143
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label lblOTANotificationPhonenumber 
            Caption         =   "Handynummer"
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame fraAdditionalInformation 
         Caption         =   "Zusatzinformationen"
         Height          =   3255
         Left            =   -69120
         TabIndex        =   156
         Top             =   3720
         Width           =   5655
         Begin VB.Label lblAutoGeneratedRemarks 
            Caption         =   "Momentane Zeit: 9.9.2002 12:35:42"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   148
            Top             =   360
            Width           =   5415
         End
         Begin VB.Label lblAutoGeneratedRemarks 
            Caption         =   "Wartezeit zwischen Mitteilungen: 5 Minuten"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   155
            Top             =   2880
            Width           =   5415
         End
         Begin VB.Label lblAutoGeneratedRemarks 
            Caption         =   "SMS Typ:Text SMS"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   151
            Top             =   1440
            Width           =   5415
         End
         Begin VB.Label lblAutoGeneratedRemarks 
            Caption         =   "Enddatum: 11.9.2002 16:00:00"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   154
            Top             =   2520
            Width           =   5415
         End
         Begin VB.Label lblAutoGeneratedRemarks 
            Caption         =   "Anzahl Empfänger: 3"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   149
            Top             =   720
            Width           =   5415
         End
         Begin VB.Label lblAutoGeneratedRemarks 
            Caption         =   "Anzahl SMS: 300"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   150
            Top             =   1080
            Width           =   5415
         End
         Begin VB.Label lblAutoGeneratedRemarks 
            Caption         =   "Versandart: Periodisch"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   152
            Top             =   1800
            Width           =   5415
         End
         Begin VB.Label lblAutoGeneratedRemarks 
            Caption         =   "Startdatum: 5.9.2002 16:00:00"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   153
            Top             =   2160
            Width           =   5415
         End
      End
      Begin VB.Frame fraJobRemarks 
         Caption         =   "Bemerkungen"
         Height          =   6375
         Left            =   -74880
         TabIndex        =   140
         Top             =   600
         Width           =   5655
         Begin VB.CommandButton cmdViewThe 
            Caption         =   "View the work reports"
            Height          =   675
            Left            =   2400
            TabIndex        =   167
            Top             =   5100
            Width           =   3135
         End
         Begin VB.CommandButton cmdSaveCurrentJobRemarksAsDefault 
            Caption         =   "Momentane Einstellungen als Standard speichern"
            Height          =   615
            Left            =   2400
            TabIndex        =   139
            Top             =   4440
            Width           =   3135
         End
         Begin VB.TextBox txtJobRemarksMemo 
            Height          =   1815
            Left            =   2400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   138
            Top             =   2520
            Width           =   3135
         End
         Begin VB.ComboBox cboJobRemarks 
            Height          =   315
            Index           =   1
            Left            =   2400
            TabIndex        =   126
            Top             =   360
            Width           =   3135
         End
         Begin VB.ComboBox cboJobRemarks 
            Height          =   315
            Index           =   2
            Left            =   2400
            TabIndex        =   128
            Top             =   720
            Width           =   3135
         End
         Begin VB.ComboBox cboJobRemarks 
            Height          =   315
            Index           =   3
            Left            =   2400
            TabIndex        =   130
            Top             =   1080
            Width           =   3135
         End
         Begin VB.ComboBox cboJobRemarks 
            Height          =   315
            Index           =   4
            Left            =   2400
            TabIndex        =   132
            Top             =   1440
            Width           =   3135
         End
         Begin VB.ComboBox cboJobRemarks 
            Height          =   315
            Index           =   5
            Left            =   2400
            TabIndex        =   134
            Top             =   1800
            Width           =   3135
         End
         Begin VB.ComboBox cboJobRemarks 
            Height          =   315
            Index           =   6
            Left            =   2400
            TabIndex        =   136
            Top             =   2160
            Width           =   3135
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   125
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   127
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   129
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   131
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   133
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   135
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label lblJobRemarksMemo 
            Caption         =   "lblJobRemarksMemo"
            Height          =   255
            Left            =   120
            TabIndex        =   137
            Top             =   2520
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdEditWAPPushSMSPicture 
         Caption         =   "Change Size / Edit"
         Height          =   375
         Left            =   -66120
         TabIndex        =   119
         Top             =   440
         Width           =   2295
      End
      Begin VB.TextBox txtWAPPushSMSPicturePath 
         Height          =   285
         Left            =   -72360
         TabIndex        =   117
         Top             =   480
         Width           =   5535
      End
      Begin VB.CommandButton cmdSelectWAPPushSMSPicture 
         Caption         =   "..."
         Height          =   300
         Left            =   -66720
         OLEDropMode     =   1  'Manual
         TabIndex        =   118
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtWAPPushSMSDescription 
         Height          =   285
         Left            =   -72360
         TabIndex        =   121
         Top             =   960
         Width           =   5535
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   -74160
         Top             =   4680
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   0
         Left            =   -71280
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   1
         Left            =   -70800
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   2
         Left            =   -70320
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   3
         Left            =   -69360
         Top             =   5160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   4
         Left            =   -71280
         Top             =   4680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   5
         Left            =   -70800
         Top             =   4680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   6
         Left            =   -69840
         Top             =   4680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   7
         Left            =   -70320
         Top             =   4680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   8
         Left            =   -69840
         Top             =   5160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   9
         Left            =   -70320
         Top             =   5160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   10
         Left            =   -69840
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   11
         Left            =   -71280
         Top             =   5160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   12
         Left            =   -69360
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   13
         Left            =   -70800
         Top             =   5160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   14
         Left            =   -69360
         Top             =   5640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   15
         Left            =   -71280
         Top             =   5640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   16
         Left            =   -69360
         Top             =   4680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   17
         Left            =   -70800
         Top             =   5640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   18
         Left            =   -69840
         Top             =   5640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   19
         Left            =   -70320
         Top             =   5640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckRandomLogo 
         Index           =   20
         Left            =   -68880
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsckLogoControl 
         Left            =   -68400
         Top             =   5520
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lblCurrentBalance 
         Caption         =   "WAIT! Checking your account status...."
         Height          =   315
         Left            =   180
         TabIndex        =   166
         Top             =   600
         Width           =   5475
      End
      Begin VB.Label lblPathPicture 
         Caption         =   "Path Picture (*.bmp)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   82
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblCounterPictureMessage 
         Height          =   255
         Left            =   -74760
         TabIndex        =   159
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Image imgPreviewPictureMessage 
         Height          =   735
         Left            =   -70920
         Top             =   900
         Width           =   2415
      End
      Begin VB.Label lblPathRingtone 
         Caption         =   "Path Ringtone (*.ott; *.txt)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   79
         Top             =   480
         Width           =   2655
      End
      Begin VB.Image imgPreviewOperatorLogo 
         Height          =   375
         Left            =   -74880
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblPathLogo 
         Caption         =   "Path Logo (*.bmp)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   61
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblUserkey 
         Caption         =   "Userkey:"
         Height          =   255
         Left            =   8760
         TabIndex        =   0
         Top             =   1020
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   6960
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblOriginator 
         Caption         =   "Sender:"
         Height          =   255
         Left            =   8760
         TabIndex        =   4
         Top             =   660
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblVCardPhoneNumber 
         Caption         =   "Phonenumber:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   89
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblVCardName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   87
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblCountOtherMessages 
         Caption         =   "Count Other Messages"
         Height          =   255
         Left            =   -72480
         TabIndex        =   106
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblCountEMailMessages 
         Caption         =   "Count EMail Messages"
         Height          =   255
         Left            =   -72480
         TabIndex        =   103
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblCountFaxMessages 
         Caption         =   "Count Fax Messages"
         Height          =   255
         Left            =   -72480
         TabIndex        =   100
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblCountVoiceMessages 
         Caption         =   "Count Voice Messages"
         Height          =   255
         Left            =   -72480
         TabIndex        =   97
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblXSer 
         Caption         =   "XSer"
         Height          =   255
         Left            =   -74880
         TabIndex        =   94
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblMessagedata 
         Caption         =   "Messagedata"
         Height          =   255
         Left            =   -74880
         TabIndex        =   92
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblPathWAPPushPicture 
         Caption         =   "lblPathWAPPushPicture"
         Height          =   255
         Left            =   -74880
         TabIndex        =   116
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblWAPPushPictureDescription 
         Caption         =   "lblWAPPushPictureDescription"
         Height          =   255
         Left            =   -74880
         TabIndex        =   120
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblWAPPushSMSPictureDetails 
         Caption         =   "lblWAPPushSMSPictureDetails"
         Height          =   255
         Left            =   -74880
         TabIndex        =   158
         Top             =   1320
         Width           =   5775
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenDatabase 
         Caption         =   "Open Database"
      End
      Begin VB.Menu mnuOpenDatabaseWithAccess2000 
         Caption         =   "Open SMS Blaster Database with Access 2000"
      End
      Begin VB.Menu mnuSeparator01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportTextfiles 
         Caption         =   "Import Textfiles..."
      End
      Begin VB.Menu mnuImportLegacy 
         Caption         =   "Import SMS Blaster Databases..."
      End
      Begin VB.Menu mnuSeparator02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAccount 
      Caption         =   "Account"
      Visible         =   0   'False
      Begin VB.Menu mnuRegistration 
         Caption         =   "Registration"
      End
      Begin VB.Menu mnuSeparator03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuyCredits 
         Caption         =   "Buy Credits"
      End
      Begin VB.Menu mnuShowCredits 
         Caption         =   "Show Credits"
      End
      Begin VB.Menu mnuSeparator12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOriginators 
         Caption         =   "Originators..."
      End
      Begin VB.Menu mnuSeparator04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJoblog 
         Caption         =   "Joblog..."
      End
      Begin VB.Menu mnuSendlog 
         Caption         =   "Sendlog..."
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Visible         =   0   'False
      Begin VB.Menu mnuLanguage 
         Caption         =   "English"
         Index           =   1
      End
      Begin VB.Menu mnuLanguage 
         Caption         =   "Deutsch"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSeparator05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGeneralSettings 
         Caption         =   "General Settings"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Info"
      Visible         =   0   'False
      Begin VB.Menu mnuInfo001 
         Caption         =   "SMS Types and required credits"
      End
      Begin VB.Menu mnuSeparator06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfo002 
         Caption         =   "aspsms.com Website"
      End
      Begin VB.Menu mnuInfo003 
         Caption         =   "Prices"
      End
      Begin VB.Menu mnuInfo004 
         Caption         =   "Latest News"
      End
      Begin VB.Menu mnuInfo005 
         Caption         =   "Supported Networks"
      End
      Begin VB.Menu mnuInfo006 
         Caption         =   "Programmers documentation"
      End
      Begin VB.Menu mnuInfo007 
         Caption         =   "FAQ - Fragen && Antworten"
      End
      Begin VB.Menu mnuInfo008 
         Caption         =   "Support"
      End
      Begin VB.Menu mnuSeparator07 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfo009 
         Caption         =   "About us"
      End
   End
   Begin VB.Menu mnuEditTop01 
      Caption         =   "mnuEditTop01"
      Visible         =   0   'False
      Begin VB.Menu mnuMarkAll01 
         Caption         =   "mnuMarkAll01"
      End
      Begin VB.Menu mnuSeparator09 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddRecipientsFromPhonebook01 
         Caption         =   "mnuAddRecipientsFromPhonebook01"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "mnuDelete01"
      End
      Begin VB.Menu mnuSeparator7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy01 
         Caption         =   "mnuCopy01"
      End
   End
   Begin VB.Menu mnuEditTop02 
      Caption         =   "mnuEditTop02"
      Visible         =   0   'False
      Begin VB.Menu mnuMarkAll02 
         Caption         =   "mnuMarkAll02"
      End
      Begin VB.Menu mnuSeparator11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveFromRecipientList02 
         Caption         =   "mnuRemoveFromRecipientList02"
      End
      Begin VB.Menu mnuClearList02 
         Caption         =   "mnuClearList02"
      End
      Begin VB.Menu mnuDuplicates02 
         Caption         =   "mnuDuplicates02"
      End
      Begin VB.Menu mnuSeparator10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy02 
         Caption         =   "mnuCopy02"
      End
   End
End
Attribute VB_Name = "frmSMSMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlCurrentPhoneBookID As Long
Dim mlCurrentTemplateID As Long

Dim msControlFileNumberOfRandomLogos As String
Dim mlNumberOfRandomLogos As Long
Dim msReceivedLogoData(0 To 20) As String

Public Sub AdjustLanguageSettings(nLanguage As Integer)
        '<EhHeader>
        On Error GoTo AdjustLanguageSettings_Err
        '</EhHeader>
        Dim i As Integer
        Dim sCaption As String
        Dim sTemp As String
        Dim sApplicationPlaceHolder As String
        Dim sNumberOfLogosPlaceHolder As String

100     VersionSpecificAction 1, nLanguage
102     frmSMSMain.caption = gsApplicationName

    '    TabMain.TabCaption(0) = LoadLanguageSpecificString(nLanguage, 90)
    '    TabMain.TabCaption(1) = LoadLanguageSpecificString(nLanguage, 91)
    '    TabMain.TabCaption(2) = LoadLanguageSpecificString(nLanguage, 92)
    '    TabMain.TabCaption(3) = LoadLanguageSpecificString(nLanguage, 93)
    '    TabMain.TabCaption(4) = LoadLanguageSpecificString(nLanguage, 94)
    '    TabMain.TabCaption(5) = LoadLanguageSpecificString(nLanguage, 95)
    '    TabMain.TabCaption(6) = LoadLanguageSpecificString(nLanguage, 96)
    '    TabMain.TabCaption(7) = LoadLanguageSpecificString(nLanguage, 97)
    '    VersionSpecificAction 24, nLanguage
    '    TabMain.TabCaption(9) = LoadLanguageSpecificString(nLanguage, 480)
    '    TabMain.TabCaption(10) = LoadLanguageSpecificString(nLanguage, 478)

104     lblUserkey.caption = LoadLanguageSpecificString(nLanguage, 99)
106     lblPassword.caption = LoadLanguageSpecificString(nLanguage, 100)
108     lblOriginator.caption = LoadLanguageSpecificString(nLanguage, 101)

110     fraSMSType.caption = LoadLanguageSpecificString(nLanguage, 102)
112     optSMSType(0).caption = LoadLanguageSpecificString(nLanguage, 103)
114     optSMSType(1).caption = LoadLanguageSpecificString(nLanguage, 104)
116     optSMSType(2).caption = LoadLanguageSpecificString(nLanguage, 105)
118     optSMSType(3).caption = LoadLanguageSpecificString(nLanguage, 106)
120     optSMSType(4).caption = LoadLanguageSpecificString(nLanguage, 107)
122     optSMSType(5).caption = LoadLanguageSpecificString(nLanguage, 108)
124     optSMSType(6).caption = LoadLanguageSpecificString(nLanguage, 480)
126     optSMSType(7).caption = LoadLanguageSpecificString(nLanguage, 478)
128     optSMSType(8).caption = LoadLanguageSpecificString(nLanguage, 109)
    
130     optSMSType(1).Enabled = False
132     optSMSType(2).Enabled = False
134     optSMSType(3).Enabled = False
136     optSMSType(4).Enabled = False
138     optSMSType(5).Enabled = False
140     optSMSType(6).Enabled = False
142     optSMSType(7).Enabled = False
144     optSMSType(8).Enabled = False

146     VersionSpecificAction 5, nLanguage

       ' fraOptions.Caption = LoadLanguageSpecificString(nLanguage, 112)

       ' cmdDeferredDeliveryTime.Caption = LoadLanguageSpecificString(nLanguage, 116)

148     fraCurrentJob.caption = "Work List:" 'LoadLanguageSpecificString(nLanguage, 117)
150     fraCurrentRecipients.caption = LoadLanguageSpecificString(nLanguage, 118)
152     lblEnterRecipientManually.caption = LoadLanguageSpecificString(nLanguage, 119)
154     lblCurrentRecipients.caption = LoadLanguageSpecificString(nLanguage, 120)

156     grdRecipients.TextMatrix(0, 0) = LoadLanguageSpecificString(nLanguage, 196)
158     grdRecipients.TextMatrix(0, 1) = LoadLanguageSpecificString(nLanguage, 197)

       ' cmdAddToRecipientList.Caption = LoadLanguageSpecificString(nLanguage, 121)
160     cmdImportRecipientsFromFile.caption = LoadLanguageSpecificString(nLanguage, 122)
162     mnuImportTextfiles.caption = LoadLanguageSpecificString(nLanguage, 122)

        'cmdRemove.Caption = LoadLanguageSpecificString(nLanguage, 123)
164     mnuRemoveFromRecipientList02.caption = LoadLanguageSpecificString(nLanguage, 123)
166     cmdClearList.caption = LoadLanguageSpecificString(nLanguage, 124)
168     mnuClearList02.caption = LoadLanguageSpecificString(nLanguage, 124)
170     cmdDuplicates.caption = LoadLanguageSpecificString(nLanguage, 460)
172     mnuDuplicates02.caption = LoadLanguageSpecificString(nLanguage, 460)
174     mnuAddRecipientsFromPhonebook01.caption = LoadLanguageSpecificString(nLanguage, 661)
176     cmdAddRecipientsFromPhonebook.caption = "<<"

178     mnuMarkAll01.caption = LoadLanguageSpecificString(nLanguage, 643)
180     mnuMarkAll02.caption = LoadLanguageSpecificString(nLanguage, 643)

       ' fraPhonebook.Caption = LoadLanguageSpecificString(nLanguage, 125)
182     lblPhonebookname.caption = LoadLanguageSpecificString(nLanguage, 126)
184     lblPhonebookPhonenumber.caption = LoadLanguageSpecificString(nLanguage, 127)

186     For i = 1 To 3
188         lblPhonebookVariableField(i).caption = GetPhonebookVariableFieldFromDatabase(gnLanguage, i)
        Next

190     lblFilter.caption = LoadLanguageSpecificString(nLanguage, 216)
       ' cmdAddNew.Caption = LoadLanguageSpecificString(nLanguage, 128)
192     cmdDelete.caption = LoadLanguageSpecificString(nLanguage, 129)
194     mnuDelete.caption = LoadLanguageSpecificString(nLanguage, 129)
        'cmdUpdate.Caption = LoadLanguageSpecificString(nLanguage, 130)

196     mnuCopy01.caption = LoadLanguageSpecificString(nLanguage, 625)
198     mnuCopy02.caption = LoadLanguageSpecificString(nLanguage, 625)

200     chkFlashingSMS.caption = LoadLanguageSpecificString(nLanguage, 132)
202     chkBlinkingSMS.caption = LoadLanguageSpecificString(nLanguage, 133)
204     lblSMSPartsNote.caption = LoadLanguageSpecificString(nLanguage, 223) & " " & LoadLanguageSpecificString(nLanguage, 224) & vbCrLf & LoadLanguageSpecificString(nLanguage, 225) & " " & LoadLanguageSpecificString(nLanguage, 226) & " " & LoadLanguageSpecificString(nLanguage, 227)
                          
206     cboPlaceHolder.Clear
208     cboPlaceHolder.AddItem LoadLanguageSpecificString(nLanguage, 234)
210     cboPlaceHolder.AddItem LoadLanguageSpecificString(nLanguage, 235)

212     For i = 1 To 3
214         cboPlaceHolder.AddItem "<" & GetPhonebookVariableFieldFromDatabase(gnLanguage, i) & ">"
        Next

216     cboPlaceHolder.ListIndex = 0

218     cmdInsertPlaceHolder.caption = LoadLanguageSpecificString(nLanguage, 228)
220     cmdPreviewSMSJob.caption = LoadLanguageSpecificString(nLanguage, 229)

222     chkReplaceMessage.caption = LoadLanguageSpecificString(nLanguage, 260)
224     chkReplaceMessageUnicode.caption = LoadLanguageSpecificString(nLanguage, 260)

226     lblUCSNote.caption = LoadLanguageSpecificString(nLanguage, 485)

228     fraCurrentMessage.caption = LoadLanguageSpecificString(nLanguage, 220) '"Current message"
230     fraCurrentMessageUnicode.caption = LoadLanguageSpecificString(nLanguage, 220) '"Current message"
232     fraTemplates.caption = LoadLanguageSpecificString(nLanguage, 230) '"Template"
234     cmdTemplateAddnew.caption = LoadLanguageSpecificString(nLanguage, 231) '"Add to templates"
236     cmdTemplateSave.caption = LoadLanguageSpecificString(nLanguage, 232) '"Save"
238     cmdTemplateDelete.caption = LoadLanguageSpecificString(nLanguage, 233) '"Delete"
240     lblCharsLeftTemplate.caption = ""
242     lblNeededPartsTemplate.caption = ""

244     lblPathLogo.caption = LoadLanguageSpecificString(nLanguage, 134)
246     cmdEditLogo.caption = LoadLanguageSpecificString(nLanguage, 136)
248     fraNetworkSettings.caption = LoadLanguageSpecificString(nLanguage, 137)
250     lblMCC.caption = LoadLanguageSpecificString(nLanguage, 138)
252     lblMNC.caption = LoadLanguageSpecificString(nLanguage, 139)
254     lblCountry.caption = LoadLanguageSpecificString(nLanguage, 140)
256     lblOperator.caption = LoadLanguageSpecificString(nLanguage, 141)

258     VersionSpecificAction 45, , , sNumberOfLogosPlaceHolder
260     sTemp = LoadLanguageSpecificString(nLanguage, 173)
262     sTemp = Replace(sTemp, gcPlaceHolder, sNumberOfLogosPlaceHolder)

264     fraRandomLogos.caption = sTemp
266     VersionSpecificAction 46

268     chkRandomLogo.caption = LoadLanguageSpecificString(nLanguage, 175)
270     cmdRandomLogoRefresh.caption = LoadLanguageSpecificString(nLanguage, 176)
272     fraSelectedRandomLogo.caption = LoadLanguageSpecificString(nLanguage, 177)

274     lblPathRingtone.caption = LoadLanguageSpecificString(nLanguage, 142)

276     lblPathPicture.caption = LoadLanguageSpecificString(nLanguage, 143)
278     cmdEditPictureMessage.caption = LoadLanguageSpecificString(nLanguage, 144)

280     lblVCardName.caption = LoadLanguageSpecificString(nLanguage, 145)
282     lblVCardPhoneNumber.caption = LoadLanguageSpecificString(nLanguage, 146)
284     chkVCardBlinkingSMS.caption = LoadLanguageSpecificString(nLanguage, 147)

286     lblPathWAPPushPicture.caption = LoadLanguageSpecificString(nLanguage, 714)
288     lblWAPPushPictureDescription.caption = LoadLanguageSpecificString(nLanguage, 716)
290     cmdEditWAPPushSMSPicture.caption = LoadLanguageSpecificString(nLanguage, 711)
292     cmdHelpWAPPush.caption = LoadLanguageSpecificString(nLanguage, 712)
294     cmdTipsAndTricksWAPPush.caption = LoadLanguageSpecificString(nLanguage, 713)

296     lblMessagedata.caption = LoadLanguageSpecificString(nLanguage, 148)
298     lblXSer.caption = LoadLanguageSpecificString(nLanguage, 149)

300     chkVoiceMessages.caption = LoadLanguageSpecificString(nLanguage, 150)
302     lblCountVoiceMessages.caption = LoadLanguageSpecificString(nLanguage, 151)
304     chkFaxMessages.caption = LoadLanguageSpecificString(nLanguage, 152)
306     lblCountFaxMessages.caption = LoadLanguageSpecificString(nLanguage, 153)
308     chkEmailMessages.caption = LoadLanguageSpecificString(nLanguage, 154)
310     lblCountEMailMessages.caption = LoadLanguageSpecificString(nLanguage, 155)
312     chkOtherMessages.caption = LoadLanguageSpecificString(nLanguage, 156)
314     lblCountOtherMessages.caption = LoadLanguageSpecificString(nLanguage, 157)
316     chkStoreMessage.caption = LoadLanguageSpecificString(nLanguage, 158)
318     chkMessageWaitingIndicationBlinkingSMS.caption = LoadLanguageSpecificString(nLanguage, 159)

320     For i = optSMSType.LBound To optSMSType.UBound

322         If optSMSType(i).Value = True Then
324             cmdSend.caption = LoadLanguageSpecificString(nLanguage, 171) & " " & SelectedSMSType(nLanguage)
                Exit For
            End If

        Next

326     For i = 1 To gcnNumberOfJobRemarkFields
328         lblJobRemarks(i).caption = GetJobRemarksFieldFromDatabase(nLanguage, i)
        Next

330     lblOTANotificationPhonenumber.caption = LoadLanguageSpecificString(nLanguage, 322)
332     fraAdditionalInformation.caption = LoadLanguageSpecificString(nLanguage, 580)
334     fraJobRemarks.caption = LoadLanguageSpecificString(nLanguage, 581)
336     lblJobRemarksMemo.caption = LoadLanguageSpecificString(nLanguage, 581)
338     cmdSaveCurrentJobRemarksAsDefault.caption = LoadLanguageSpecificString(nLanguage, 582)

340     txtRecipient.toolTipText = LoadLanguageSpecificString(nLanguage, 255)
342     txtPhonebookPhoneNumber.toolTipText = LoadLanguageSpecificString(nLanguage, 255)
344     txtRecipientDeliveryNotification.toolTipText = LoadLanguageSpecificString(nLanguage, 255)

346     txtSMS_Change
348     RecipientDeliveryNotificationEventChange nLanguage
350     txtUnicode_Change
352     txtPictureMessageText_Change
354     txtFilter_Change

356     UpdateJobList nLanguage
        '<EhFooter>
        Exit Sub

AdjustLanguageSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.AdjustLanguageSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function DuplicatesFoundWithinCurrentRecipients(lDuplicates As Long) As Boolean
        '<EhHeader>
        On Error GoTo DuplicatesFoundWithinCurrentRecipients_Err
        '</EhHeader>
        Dim i As Long
        Dim sBackup As String
        Dim lIndex As Long

100     If grdRecipients.Rows <= 1 Then
102         DuplicatesFoundWithinCurrentRecipients = False
            Exit Function
        End If

104     ReDim tOrigPhoneBookEntry(1 To grdRecipients.Rows - 1) As PhoneBookEntry
106     ReDim tDestPhoneBookEntry(1 To grdRecipients.Rows - 1) As PhoneBookEntry

108     Screen.MousePointer = vbHourglass

        'Fill Array
110     For i = LBound(tOrigPhoneBookEntry) To UBound(tOrigPhoneBookEntry)
112         tOrigPhoneBookEntry(i).sNumber = grdRecipients.TextMatrix(i, 1)
        Next

        'Sort Array by PhoneNumber
114     QuickSortPhoneBookEntryArrayByPhoneNumber tOrigPhoneBookEntry, LBound(tOrigPhoneBookEntry), UBound(tOrigPhoneBookEntry)

116     lIndex = LBound(tOrigPhoneBookEntry)

118     For i = LBound(tOrigPhoneBookEntry) To UBound(tOrigPhoneBookEntry)

120         If tOrigPhoneBookEntry(i).sNumber <> sBackup Then
122             tDestPhoneBookEntry(lIndex) = tOrigPhoneBookEntry(i)
124             lIndex = lIndex + 1
            Else
126             lDuplicates = lDuplicates + 1
            End If

128         sBackup = tOrigPhoneBookEntry(i).sNumber
        Next

130     Screen.MousePointer = vbDefault

132     If lDuplicates > 0 Then
134         DuplicatesFoundWithinCurrentRecipients = True
        Else
136         DuplicatesFoundWithinCurrentRecipients = False
        End If

        '<EhFooter>
        Exit Function

DuplicatesFoundWithinCurrentRecipients_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.DuplicatesFoundWithinCurrentRecipients " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub FormLoadWithoutSubClassing()
        '<EhHeader>
        On Error GoTo FormLoadWithoutSubClassing_Err
        '</EhHeader>
        Dim sCommand As String
100     sCommand = Command

102     If sCommand <> "" Then
104         If InStr(1, sCommand, "\") >= 1 Then
                'Ok, probabely valid path
106             gsPathAndDatabaseName = Command
            Else
108             gsPathAndDatabaseName = App.Path & "\" & sCommand
            End If
        End If

110     CenterForm Me
112     InitApp
114     TabMain.Tab = 0
        '<EhFooter>
        Exit Sub

FormLoadWithoutSubClassing_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.FormLoadWithoutSubClassing " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub RecipientDeliveryNotificationEventChange(nLanguage As Integer)
        '<EhHeader>
        On Error GoTo RecipientDeliveryNotificationEventChange_Err
        '</EhHeader>
        Dim sMessage As String
        Dim sRecipient As String

100     If Trim(txtRecipientDeliveryNotification.Text) = "" Then
102         sRecipient = "-"
        Else
104         sRecipient = txtRecipientDeliveryNotification.Text
        End If

106     sMessage = LoadLanguageSpecificString(gnLanguage, 315) & " " & sRecipient & LoadLanguageSpecificString(gnLanguage, 316)
108     chkUseOTADeliveryNotifications.caption = LoadLanguageSpecificString(nLanguage, 313)
110     chkDeliveryNotificationBuffered.caption = sMessage & LoadLanguageSpecificString(nLanguage, 317)
112     chkDeliveryNotificationDelivered.caption = sMessage & LoadLanguageSpecificString(nLanguage, 318)
114     chkDeliveryNotificationNotDelivered.caption = sMessage & LoadLanguageSpecificString(nLanguage, 319)
116     fraDeliveryNotificationsPerSMS.caption = LoadLanguageSpecificString(nLanguage, 311)

        '<EhFooter>
        Exit Sub

RecipientDeliveryNotificationEventChange_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.RecipientDeliveryNotificationEventChange " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub ResizerResizeControl(ctlResize1 As Control, _
                         ctlReference As Control, _
                         nMode As Integer, _
                         Optional ctlResize2 As Control)
        '<EhHeader>
        On Error GoTo ResizerResizeControl_Err
        '</EhHeader>
        Dim lTemp As Long
        Const gcResizerRightAlignToContainerConstantWidth = 1
        Const gcResizerRightAlignToNonContainerVariableWidth = 2
        Const gcResizerResizeToRightConstantLeft = 3

        Const gcResizerBottomAlignToContainerConstantHeight = 1
        Const gcResizerBottomAlignToNonContainerVariableHeight = 2

        Const gcResizerBottomAlignConstantHeight = 4
        Const gcResizerResizeToBottomConstantTop = 5

100     Select Case nMode

            Case gcResizerRightAlignToContainerConstantWidth
102             lTemp = ctlReference.Width - ctlResize1.Width - 120
104             ctlResize1.Left = IIf(lTemp > 0, lTemp, 0)
  
106         Case gcResizerRightAlignToNonContainerVariableWidth
108             lTemp = ctlReference.Left - ctlResize1.Left - 120
110             ctlResize1.Width = IIf(lTemp > 0, lTemp, 0)
    
112         Case gcResizerResizeToRightConstantLeft

114         Case gcResizerBottomAlignConstantHeight

116         Case gcResizerResizeToBottomConstantTop

        End Select

        '<EhFooter>
        Exit Sub

ResizerResizeControl_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.ResizerResizeControl " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub SetCorrectTopRowWithinPhonebook(lID As Long)
        '<EhHeader>
        On Error GoTo SetCorrectTopRowWithinPhonebook_Err
        '</EhHeader>
        Dim lMaxRows As Long
        Dim lCurrentRow As Long
        Dim lHeightOfAllRowsAbove As Long

100     lMaxRows = grdPhonebook.Rows - 1

102     DoEvents
104     lCurrentRow = 0

106     Do While lCurrentRow < lMaxRows And Val(grdPhonebook.TextMatrix(lCurrentRow, 0)) <> lID
108         lHeightOfAllRowsAbove = lHeightOfAllRowsAbove + grdPhonebook.CellTop
110         lCurrentRow = lCurrentRow + 1

112         DoEvents
        Loop

114     If grdPhonebook.Height > lHeightOfAllRowsAbove Then
116         grdPhonebook.Row = lCurrentRow
        Else
118         grdPhonebook.Row = lCurrentRow
120         grdPhonebook.TopRow = lCurrentRow
        End If

        '<EhFooter>
        Exit Sub

SetCorrectTopRowWithinPhonebook_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.SetCorrectTopRowWithinPhonebook " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function TextSMSLengthInfo(sInput As String, bBlinkingSMS As Boolean, nNumberOfChars As Integer, nNumberOfConcatenationParts As Integer)
        '<EhHeader>
        On Error GoTo TextSMSLengthInfo_Err
        '</EhHeader>
        Dim sTemp As String
        'VBCRLF is replaced by <Space>, because it only uses 1 character
100     sTemp = Replace(sInput, vbCrLf, " ")

102     If bBlinkingSMS Then    'Characters left is depending on using BlinkingSMS Feature or not

            'BlinkingSMS used
            'Set BlinkTag if not yet done
104         If InStr(1, UCase(sTemp), "<BLINK>") = 0 And InStr(1, UCase(sTemp), "</BLINK>") = 0 Then
106             sTemp = "<BLINK>" & sTemp
            End If

            '<BLINK> is replaced by <Space>, because it only uses 1 character
108         sTemp = Replace(sTemp, "<BLINK>", " ", , , vbTextCompare)
110         sTemp = Replace(sTemp, "</BLINK>", " ", , , vbTextCompare)
  
        End If

112     nNumberOfChars = Len(sTemp)
114     nNumberOfConcatenationParts = NumberOfConcatenationParts(Len(sTemp), bBlinkingSMS)

        '<EhFooter>
        Exit Function

TextSMSLengthInfo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.TextSMSLengthInfo " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub chkUseOTADeliveryNotifications_Click()
        '<EhHeader>
        On Error GoTo chkUseOTADeliveryNotifications_Click_Err
        '</EhHeader>
100     txtRecipientDeliveryNotification.Enabled = Not txtRecipientDeliveryNotification.Enabled
102     chkDeliveryNotificationBuffered.Enabled = Not chkDeliveryNotificationBuffered.Enabled
104     chkDeliveryNotificationDelivered.Enabled = Not chkDeliveryNotificationDelivered.Enabled
106     chkDeliveryNotificationNotDelivered.Enabled = Not chkDeliveryNotificationNotDelivered.Enabled
        '<EhFooter>
        Exit Sub

chkUseOTADeliveryNotifications_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.chkUseOTADeliveryNotifications_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAddRecipientsFromPhonebook_Click()
        '<EhHeader>
        On Error GoTo cmdAddRecipientsFromPhonebook_Click_Err
        '</EhHeader>
        Dim lStart As Long
        Dim lEnd As Long
        Dim i As Long

100     If grdPhonebook.Rows <= 1 Then
            Exit Sub
        End If

102     If grdPhonebook.RowSel < grdPhonebook.Row Then
104         lStart = grdPhonebook.RowSel
106         lEnd = grdPhonebook.Row
        Else
108         lStart = grdPhonebook.Row
110         lEnd = grdPhonebook.RowSel
        End If

112     Screen.MousePointer = vbHourglass

114     For i = lStart To lEnd
116         grdRecipients.AddItem grdPhonebook.TextMatrix(i, 1) & Chr$(9) & grdPhonebook.TextMatrix(i, 2) & Chr$(9) & grdPhonebook.TextMatrix(i, 3) & Chr$(9) & grdPhonebook.TextMatrix(i, 4) & Chr$(9) & grdPhonebook.TextMatrix(i, 5)
        Next

118     Screen.MousePointer = vbDefault

120     UpdateJobList gnLanguage
        '<EhFooter>
        Exit Sub

cmdAddRecipientsFromPhonebook_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdAddRecipientsFromPhonebook_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkBlinkingSMS_Click()
        '<EhHeader>
        On Error GoTo chkBlinkingSMS_Click_Err
        '</EhHeader>
100     txtSMS_Change
        '<EhFooter>
        Exit Sub

chkBlinkingSMS_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.chkBlinkingSMS_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAddToRecipientList_Click()
        '<EhHeader>
        On Error GoTo cmdAddToRecipientList_Click_Err
        '</EhHeader>
        Dim i As Integer

100     If txtRecipient.Text <> "" Then
102         grdRecipients.AddItem "" & Chr$(9) & txtRecipient.Text
        End If

104     UpdateJobList gnLanguage

        '<EhFooter>
        Exit Sub

cmdAddToRecipientList_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdAddToRecipientList_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdDeferredDeliveryTime_Click()
        '<EhHeader>
        On Error GoTo cmdDeferredDeliveryTime_Click_Err
        '</EhHeader>
100     frmDeferredDeliveryTime.Show vbModal, Me
        '<EhFooter>
        Exit Sub

cmdDeferredDeliveryTime_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdDeferredDeliveryTime_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdDuplicates_Click()
        '<EhHeader>
        On Error GoTo cmdDuplicates_Click_Err
        '</EhHeader>
        Dim i As Long
        Dim sBackup As String
        Dim lIndex As Long
        Dim lDuplicates As Long
        Dim sTemp As String
        Dim sMessage As String
        Dim sInfoDuplicates As String
        Dim nShowDuplicates As Integer
        Dim nAnswer As Integer
        Const cMaxDuplicatesToShow = 20

100     If grdRecipients.Rows <= 1 Then Exit Sub

102     ReDim tOrigPhoneBookEntry(1 To grdRecipients.Rows - 1) As PhoneBookEntry
104     ReDim tDestPhoneBookEntry(1 To grdRecipients.Rows - 1) As PhoneBookEntry
106     ReDim tFoundDuplicatesPhoneBookEntry(1 To cMaxDuplicatesToShow) As PhoneBookEntry

108     Screen.MousePointer = vbHourglass

        'Fill Array
110     For i = LBound(tOrigPhoneBookEntry) To UBound(tOrigPhoneBookEntry)
112         tOrigPhoneBookEntry(i).sName = grdRecipients.TextMatrix(i, 0)
114         tOrigPhoneBookEntry(i).sNumber = grdRecipients.TextMatrix(i, 1)
116         tOrigPhoneBookEntry(i).sVariable1 = grdRecipients.TextMatrix(i, 2)
118         tOrigPhoneBookEntry(i).sVariable2 = grdRecipients.TextMatrix(i, 3)
120         tOrigPhoneBookEntry(i).sVariable3 = grdRecipients.TextMatrix(i, 4)
        Next

        'Sort Array by PhoneNumber
122     QuickSortPhoneBookEntryArrayByPhoneNumber tOrigPhoneBookEntry, LBound(tOrigPhoneBookEntry), UBound(tOrigPhoneBookEntry)

        'Remove Duplicates
124     lIndex = LBound(tOrigPhoneBookEntry)

126     For i = LBound(tOrigPhoneBookEntry) To UBound(tOrigPhoneBookEntry)

128         If tOrigPhoneBookEntry(i).sNumber <> sBackup Then
130             tDestPhoneBookEntry(lIndex) = tOrigPhoneBookEntry(i)
132             lIndex = lIndex + 1
            Else
134             lDuplicates = lDuplicates + 1

136             If lDuplicates <= cMaxDuplicatesToShow Then
138                 tFoundDuplicatesPhoneBookEntry(lDuplicates) = tOrigPhoneBookEntry(i)
                End If
            End If

140         sBackup = tOrigPhoneBookEntry(i).sNumber
        Next

142     Screen.MousePointer = vbDefault

144     If lDuplicates > 0 Then
146         If lDuplicates > cMaxDuplicatesToShow Then
148             nShowDuplicates = cMaxDuplicatesToShow
            Else
150             nShowDuplicates = lDuplicates
            End If
  
152         For i = 1 To nShowDuplicates

154             If tFoundDuplicatesPhoneBookEntry(i).sName = "" Then
156                 sInfoDuplicates = sInfoDuplicates & tFoundDuplicatesPhoneBookEntry(i).sNumber & vbCrLf
                Else
158                 sInfoDuplicates = sInfoDuplicates & tFoundDuplicatesPhoneBookEntry(i).sName & " / " & tFoundDuplicatesPhoneBookEntry(i).sNumber & vbCrLf
                End If

            Next

160         If lDuplicates > nShowDuplicates Then
162             sInfoDuplicates = sInfoDuplicates & LoadLanguageSpecificString(gnLanguage, 459) & vbCrLf
            End If
  
164         If lDuplicates = 1 Then
166             sMessage = LoadLanguageSpecificString(gnLanguage, 453) & vbCrLf & vbCrLf & sInfoDuplicates & vbCrLf & LoadLanguageSpecificString(gnLanguage, 454)
            Else
168             sMessage = LoadLanguageSpecificString(gnLanguage, 455) & " " & Trim(Str$(lDuplicates)) & " " & LoadLanguageSpecificString(gnLanguage, 456) & vbCrLf & vbCrLf & sInfoDuplicates & vbCrLf & LoadLanguageSpecificString(gnLanguage, 457)
            End If

170         nAnswer = MsgBox(sMessage, vbQuestion Or vbOKCancel, gsApplicationName)

172         If nAnswer = vbOK Then
                'Proceed
            Else
                Exit Sub
            End If

        Else
174         sMessage = LoadLanguageSpecificString(gnLanguage, 458)
176         MsgBox sMessage, vbInformation, gsApplicationName
            Exit Sub
        End If
  
178     ReDim Preserve tDestPhoneBookEntry(LBound(tDestPhoneBookEntry) To lIndex - 1) As PhoneBookEntry

        'Sort Array by Name
180     QuickSortPhoneBookEntryArrayByName tDestPhoneBookEntry, LBound(tDestPhoneBookEntry), UBound(tDestPhoneBookEntry)

        'Clear list
182     grdRecipients.Rows = 1

        'Refill list
184     Screen.MousePointer = vbHourglass

186     For i = LBound(tDestPhoneBookEntry) To lIndex - 1
188         sTemp = tDestPhoneBookEntry(i).sName & Chr$(9) & tDestPhoneBookEntry(i).sNumber & Chr$(9) & tDestPhoneBookEntry(i).sVariable1 & Chr$(9) & tDestPhoneBookEntry(i).sVariable2 & Chr$(9) & tDestPhoneBookEntry(i).sVariable3
190         grdRecipients.AddItem sTemp
        Next

192     Screen.MousePointer = vbDefault

194     UpdateJobList gnLanguage
        '<EhFooter>
        Exit Sub

cmdDuplicates_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdDuplicates_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdEditLogo_Click()
        '<EhHeader>
        On Error GoTo cmdEditLogo_Click_Err
        '</EhHeader>
        Dim lRet As Long
100     lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", txtPathLogo.Text, vbNullString, vbNullString, vbNormalFocus)

        '<EhFooter>
        Exit Sub

cmdEditLogo_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdEditLogo_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdEditPictureMessage_Click()
        '<EhHeader>
        On Error GoTo cmdEditPictureMessage_Click_Err
        '</EhHeader>
        Dim lRet As Long
100     lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", txtPathPictureMessage.Text, vbNullString, vbNullString, vbNormalFocus)
        '<EhFooter>
        Exit Sub

cmdEditPictureMessage_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdEditPictureMessage_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdGeneralSettings_Click()
        '<EhHeader>
        On Error GoTo cmdGeneralSettings_Click_Err
        '</EhHeader>
100     frmSettings.Show vbModal, Me
        '<EhFooter>
        Exit Sub

cmdGeneralSettings_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdGeneralSettings_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdEditWAPPushSMSPicture_Click()
        '<EhHeader>
        On Error GoTo cmdEditWAPPushSMSPicture_Click_Err
        '</EhHeader>
        Dim lRet As Long
100     lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", txtWAPPushSMSPicturePath.Text, vbNullString, vbNullString, vbNormalFocus)

        '<EhFooter>
        Exit Sub

cmdEditWAPPushSMSPicture_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdEditWAPPushSMSPicture_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdHelpWAPPush_Click()
        '<EhHeader>
        On Error GoTo cmdHelpWAPPush_Click_Err
        '</EhHeader>
        Dim sTemp As String
        Dim sMessage As String
        Dim sDirectoryPlaceHolder As String
        Dim sApplicationPlaceHolder As String

100     sDirectoryPlaceHolder = LoadLanguageSpecificString(gnLanguage, 731)
102     sDirectoryPlaceHolder = Replace(sDirectoryPlaceHolder, gcPlaceHolder, App.Path & "\Images")
104     sTemp = sTemp & sDirectoryPlaceHolder

106     VersionSpecificAction 44, , , sApplicationPlaceHolder
108     sMessage = LoadLanguageSpecificString(gnLanguage, 725) & vbCrLf & vbCrLf

110     sTemp = LoadLanguageSpecificString(gnLanguage, 726)
112     sTemp = Replace(sTemp, gcPlaceHolder, sApplicationPlaceHolder)
114     sMessage = sMessage & sTemp & vbCrLf & vbCrLf

116     sTemp = LoadLanguageSpecificString(gnLanguage, 727)
118     sTemp = Replace(sTemp, gcPlaceHolder, sApplicationPlaceHolder)
120     sMessage = sMessage & sTemp & vbCrLf & vbCrLf

122     sMessage = sMessage & LoadLanguageSpecificString(gnLanguage, 728) & vbCrLf & vbCrLf
124     sMessage = sMessage & LoadLanguageSpecificString(gnLanguage, 729) & vbCrLf & vbCrLf
126     sMessage = sMessage & LoadLanguageSpecificString(gnLanguage, 730) & vbCrLf & vbCrLf

128     sDirectoryPlaceHolder = LoadLanguageSpecificString(gnLanguage, 731)
130     sDirectoryPlaceHolder = Replace(sDirectoryPlaceHolder, gcPlaceHolder, App.Path & "\Images")
132     sMessage = sMessage & sDirectoryPlaceHolder

134     MsgBox sMessage, vbInformation, gsApplicationName
        '<EhFooter>
        Exit Sub

cmdHelpWAPPush_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdHelpWAPPush_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdImportRecipientsFromFile_Click()
        '<EhHeader>
        On Error GoTo cmdImportRecipientsFromFile_Click_Err
        '</EhHeader>
100     'frmImport.Show vbModal, Me
        '<EhFooter>
        Exit Sub

cmdImportRecipientsFromFile_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdImportRecipientsFromFile_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdInsertPlaceHolder_Click()
        '<EhHeader>
        On Error GoTo cmdInsertPlaceHolder_Click_Err
        '</EhHeader>
        Dim sTemp As String
        Dim nSelStart As Integer

100     nSelStart = txtSMS.SelStart
102     sTemp = txtSMS.Text

104     sTemp = Left(sTemp, nSelStart) & cboPlaceHolder.List(cboPlaceHolder.ListIndex) & Right(sTemp, Len(sTemp) - nSelStart)

106     txtSMS.Text = sTemp
108     txtSMS.SelStart = nSelStart + Len(cboPlaceHolder.List(cboPlaceHolder.ListIndex))

110     txtSMS.SetFocus

        '<EhFooter>
        Exit Sub

cmdInsertPlaceHolder_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdInsertPlaceHolder_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdPreviewSMSJob_Click()
'        '<EhHeader>
'        On Error GoTo cmdPreviewSMSJob_Click_Err
'        '</EhHeader>
'        Dim lRecipient As Long
'        Dim sMessage As String
'        Dim sInputMessage As String
'        Dim bVariableMessageData As Boolean
'        Dim bBlinkingSMS As Boolean
'        Dim nNumberOfChars As Integer
'        Dim nNumberOfConcatenationParts As Integer
'
'100     sInputMessage = txtSMS.Text
'
'102     If chkBlinkingSMS.Value = 1 Then
'104         bBlinkingSMS = True
'        Else
'106         bBlinkingSMS = False
'        End If
'
'108     Screen.MousePointer = vbHourglass
'
'110     For lRecipient = 1 To grdRecipients.Rows - 1
'112         sMessage = CreatePersonalizedMessageData(sInputMessage, lRecipient, bVariableMessageData)
'114         TextSMSLengthInfo sMessage, bBlinkingSMS, nNumberOfChars, nNumberOfConcatenationParts
'
'116         sMessage = LoadLanguageSpecificString(gnLanguage, 218) & ": " & Trim(Str$(nNumberOfConcatenationParts) & " / " & LoadLanguageSpecificString(gnLanguage, 219) & ": " & Trim(Str$(nNumberOfChars))) & " / " & sMessage
'
'118         frmSMSJobPreview.lstPreview.AddItem sMessage
'        Next
'
'120     If (grdRecipients.Rows - 1) = 0 Then
'122         MsgBox LoadLanguageSpecificString(gnLanguage, 653), vbInformation, gsApplicationName
'        End If
'
'124     Screen.MousePointer = vbDefault
'126     frmSMSJobPreview.Show vbModal, Me
'        '<EhFooter>
'        Exit Sub
'
'cmdPreviewSMSJob_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmSMSMain.cmdPreviewSMSJob_Click " & _
'               "at line " & Erl
'        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdRandomLogoRefresh_Click()
        '<EhHeader>
        On Error GoTo cmdRandomLogoRefresh_Click_Err
        '</EhHeader>
        Dim sData As String
        Dim bytHTML() As Byte
        Dim sRessource As String
        Dim sURL As String
        Dim sHost As String
        Dim i As Integer
        Dim nLogo As Integer
        Dim sPicture As String
        Dim sFilename As String
        Dim sRandomLogoFilename As String
        Dim lNumberOfRandomLogos As Long
        Dim sControlFileNumberOfRandomLogos As String
        Dim sLogoPath As String
        Dim nRandomLogo As Integer
        Dim bSuccess As Boolean
        Static bEventOngoing As Boolean

        On Error GoTo ErrorTrap

100     If bEventOngoing Then Exit Sub
102     bEventOngoing = True

104     VersionSpecificAction 39, , , sRessource
106     VersionSpecificAction 40, , , sHost

108     sURL = "http://" & sHost & sRessource

110     With Inet1
112         .Cancel
114         sURL = "http://" & sHost & sRessource
116         sControlFileNumberOfRandomLogos = .OpenURL(sURL)
        End With

118     lNumberOfRandomLogos = ParseReceivedLogoControlData(sControlFileNumberOfRandomLogos, bSuccess)
  
120     If bSuccess Then
            'Continue
        Else
            Exit Sub
        End If

122     For nLogo = 0 To 20

124         nRandomLogo = Int(Rnd * lNumberOfRandomLogos + 1)
126         sRandomLogoFilename = Right("000" & Trim(Str(nRandomLogo)), 4) & ".bmp"

128         With Inet1
130             .Cancel
    
132             VersionSpecificAction 40, , , sHost
134             VersionSpecificAction 41, , , sLogoPath

136             Randomize Second(Now)
138             nRandomLogo = Int(Rnd * lNumberOfRandomLogos + 1)
140             sRandomLogoFilename = Right("000" & Trim(Str(nRandomLogo)), 4) & ".bmp"
142             sURL = "http://" & sHost & sLogoPath & sRandomLogoFilename
    
144             bytHTML = .OpenURL(sURL, icByteArray)
            End With
  
146         sPicture = ""

148         For i = 0 To UBound(bytHTML) - 1
150             sPicture = sPicture & Chr(bytHTML(i))
            Next
  
152         ProcessRandomLogoDisplayUpdate nLogo, sPicture
154         sFilename = App.Path & "\" & Right("0000" & Trim(Str$(nLogo)), 4) & ".bmp"
156         imgRandomLogo(nLogo).Picture = LoadPicture(sFilename)
  
158         If nLogo = 0 Then
160             imgRandomLogo_Click 0
            End If

        Next

162     bEventOngoing = False
        Exit Sub

ErrorTrap:
        Exit Sub
        '<EhFooter>
        Exit Sub

cmdRandomLogoRefresh_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdRandomLogoRefresh_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSaveCurrentJobRemarksAsDefault_Click()
        '<EhHeader>
        On Error GoTo cmdSaveCurrentJobRemarksAsDefault_Click_Err
        '</EhHeader>
        Dim i As Integer

100     For i = 1 To gcnNumberOfJobRemarkFields
102         PutSettingIntoDataBase "JobRemarksDefault" & Right("00" & Trim(Str$(i)), 2), Left(cboJobRemarks(i).Text, 50)
        Next

104     PutSettingIntoDataBase "OTANotificationSettingsMobilenumber", txtRecipientDeliveryNotification.Text
 
106     If chkUseOTADeliveryNotifications.Value = 1 Then
108         PutSettingIntoDataBase "OTANotificationSettingsUseService", True
        Else
110         PutSettingIntoDataBase "OTANotificationSettingsUseService", False
        End If

112     If chkDeliveryNotificationBuffered.Value = 1 Then
114         PutSettingIntoDataBase "OTANotificationSettingsEventBuffered", True
        Else
116         PutSettingIntoDataBase "OTANotificationSettingsEventBuffered", False
        End If

118     If chkDeliveryNotificationDelivered.Value = 1 Then
120         PutSettingIntoDataBase "OTANotificationSettingsEventDelivered", True
        Else
122         PutSettingIntoDataBase "OTANotificationSettingsEventDelivered", False
        End If

124     If chkDeliveryNotificationNotDelivered.Value = 1 Then
126         PutSettingIntoDataBase "OTANotificationSettingsEventNotDelivered", True
        Else
128         PutSettingIntoDataBase "OTANotificationSettingsEventNotDelivered", False
        End If

        '<EhFooter>
        Exit Sub

cmdSaveCurrentJobRemarksAsDefault_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdSaveCurrentJobRemarksAsDefault_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSelectWAPPushSMSPicture_Click()
        '<EhHeader>
        On Error GoTo cmdSelectWAPPushSMSPicture_Click_Err
        '</EhHeader>
        On Error GoTo ErrorCancelSelected
        Dim sTemp As String
        Dim sTempArray() As String

100     cmdDlgOpen.Filter = LoadLanguageSpecificString(gnLanguage, 715)
102     cmdDlgOpen.ShowOpen
104     txtWAPPushSMSPicturePath.Text = cmdDlgOpen.Filename
106     txtWAPPushSMSDescription.Text = ExtractFilenameWithoutExtension(txtWAPPushSMSPicturePath.Text)

108     Form_Paint

        Exit Sub
ErrorCancelSelected:
        Exit Sub

        '<EhFooter>
        Exit Sub

cmdSelectWAPPushSMSPicture_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdSelectWAPPushSMSPicture_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTemplateDelete_Click()
        '<EhHeader>
        On Error GoTo cmdTemplateDelete_Click_Err
        '</EhHeader>
        Dim lStart As Long
        Dim lEnd As Long
        Dim nAnswer As Integer
        Dim i As Long

100     If grdTemplates.RowSel < grdTemplates.Row Then
102         lStart = grdTemplates.RowSel
104         lEnd = grdTemplates.Row
        Else
106         lStart = grdTemplates.Row
108         lEnd = grdTemplates.RowSel
        End If

110     nAnswer = MsgBox(LoadLanguageSpecificString(gnLanguage, 208) & Str$(lEnd - lStart + 1) & " " & LoadLanguageSpecificString(gnLanguage, 209), vbOKCancel Or vbQuestion, gsApplicationName)

112     If nAnswer = vbOK Then

114         For i = lStart To lEnd
116             dtaTemplates.Recordset.FindFirst "ID = " & grdTemplates.TextMatrix(i, 0)
118             dtaTemplates.Recordset.Delete
            Next

120         dtaTemplates.Refresh
        End If

        '<EhFooter>
        Exit Sub

cmdTemplateDelete_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdTemplateDelete_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTemplateAddnew_Click()
        '<EhHeader>
        On Error GoTo cmdTemplateAddnew_Click_Err
        '</EhHeader>
100     dtaTemplates.Recordset.AddNew
102     dtaTemplates.Recordset("Message") = txtTemplateMessage.Text
104     dtaTemplates.Recordset.UpDate

106     dtaTemplates.Recordset.MoveLast
108     mlCurrentTemplateID = dtaTemplates.Recordset("ID")
110     dtaTemplates.Refresh

        '<EhFooter>
        Exit Sub

cmdTemplateAddnew_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdTemplateAddnew_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTemplateCopy_Click()
        '<EhHeader>
        On Error GoTo cmdTemplateCopy_Click_Err
        '</EhHeader>
100     txtSMS.Text = txtTemplateMessage.Text
        '<EhFooter>
        Exit Sub

cmdTemplateCopy_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdTemplateCopy_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTemplateSave_Click()
        '<EhHeader>
        On Error GoTo cmdTemplateSave_Click_Err
        '</EhHeader>

100     dtaTemplates.Recordset.FindFirst "ID = " & mlCurrentTemplateID

102     If dtaTemplates.Recordset.RecordCount = 0 Then
104         dtaTemplates.Recordset.AddNew
        Else
106         dtaTemplates.Recordset.Edit
        End If

108     dtaTemplates.Recordset("Message") = txtTemplateMessage.Text
110     dtaTemplates.Recordset.UpDate
112     dtaTemplates.Refresh

        '<EhFooter>
        Exit Sub

cmdTemplateSave_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdTemplateSave_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function HexByteToDez(sInput As String) As Byte
        '<EhHeader>
        On Error GoTo HexByteToDez_Err
        '</EhHeader>
        Dim sTemp As String
100     ReDim nDez(1 To 2) As Integer
        Dim i As Integer

102     sTemp = UCase(Trim(sInput))

104     If Len(sTemp) <> 2 Then
106         sTemp = Right("00" & sTemp, 2)
        End If

108     For i = 1 To 2

110         Select Case Mid(sTemp, i, 1)
  
                Case "0"
112                 nDez(i) = Val(Mid(sTemp, i, 1))
  
114             Case "1"
116                 nDez(i) = Mid(sTemp, i, 1)
  
118             Case "2"
120                 nDez(i) = Mid(sTemp, i, 1)
  
122             Case "3"
124                 nDez(i) = Mid(sTemp, i, 1)
  
126             Case "4"
128                 nDez(i) = Mid(sTemp, i, 1)
  
130             Case "5"
132                 nDez(i) = Mid(sTemp, i, 1)
  
134             Case "6"
136                 nDez(i) = Mid(sTemp, i, 1)
  
138             Case "7"
140                 nDez(i) = Mid(sTemp, i, 1)
  
142             Case "8"
144                 nDez(i) = Mid(sTemp, i, 1)
  
146             Case "9"
148                 nDez(i) = Mid(sTemp, i, 1)
  
150             Case "A"
152                 nDez(i) = 10
  
154             Case "B"
156                 nDez(i) = 11
  
158             Case "C"
160                 nDez(i) = 12
  
162             Case "D"
164                 nDez(i) = 13
  
166             Case "E"
168                 nDez(i) = 14
  
170             Case "F"
172                 nDez(i) = 15
  
174             Case Else
176                 nDez(i) = 0
  
            End Select

        Next

178     HexByteToDez = nDez(1) * 16 + nDez(2)

        '<EhFooter>
        Exit Function

HexByteToDez_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.HexByteToDez " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdTipsAndTricksWAPPush_Click()
        '<EhHeader>
        On Error GoTo cmdTipsAndTricksWAPPush_Click_Err
        '</EhHeader>
        Dim sTemp As String
        Dim sTempPlaceHolder As String
100     sTemp = sTemp & LoadLanguageSpecificString(gnLanguage, 734) & vbCrLf
102     sTemp = sTemp & LoadLanguageSpecificString(gnLanguage, 735) & vbCrLf & vbCrLf
104     sTemp = sTemp & LoadLanguageSpecificString(gnLanguage, 736) & vbCrLf & vbCrLf
106     sTemp = sTemp & LoadLanguageSpecificString(gnLanguage, 737) & vbCrLf & vbCrLf
108     sTemp = sTemp & LoadLanguageSpecificString(gnLanguage, 738)
110     MsgBox sTemp, vbInformation, gsApplicationName
        '<EhFooter>
        Exit Sub

cmdTipsAndTricksWAPPush_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdTipsAndTricksWAPPush_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, _
                             Effect As Long, _
                             Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single, _
                             State As Integer)
        '<EhHeader>
        On Error GoTo Form_OLEDragOver_Err
        '</EhHeader>
100     Me.Show
        '<EhFooter>
        Exit Sub

Form_OLEDragOver_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.Form_OLEDragOver " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Paint()
        '<EhHeader>
        On Error GoTo Form_Paint_Err
        '</EhHeader>
        Dim sTemp As String
        On Error Resume Next

100     If txtWAPPushSMSPicturePath.Text <> "" Then
102         picContainer.Visible = True
104         imgPreviewWAPPushSMSPicture.Width = 1
106         imgPreviewWAPPushSMSPicture.Height = 1
  
108         imgPreviewWAPPushSMSPicture.Refresh
110         imgPreviewWAPPushSMSPicture.Picture = LoadPicture(cmdDlgOpen.Filename)

112         sTemp = LoadLanguageSpecificString(gnLanguage, 718) & " " & Trim(Str(imgPreviewWAPPushSMSPicture.Width / Screen.TwipsPerPixelX)) & " "
114         sTemp = sTemp & LoadLanguageSpecificString(gnLanguage, 720) & " / "
116         sTemp = sTemp & LoadLanguageSpecificString(gnLanguage, 719) & " " & Trim(Str(imgPreviewWAPPushSMSPicture.Height / Screen.TwipsPerPixelY)) & " "
118         sTemp = sTemp & LoadLanguageSpecificString(gnLanguage, 720) & " / "
120         sTemp = sTemp & LoadLanguageSpecificString(gnLanguage, 721) & " " & Trim(Str(FileLen(cmdDlgOpen.Filename))) & " "
122         sTemp = sTemp & LoadLanguageSpecificString(gnLanguage, 722)

124         lblWAPPushSMSPictureDetails.caption = sTemp
        Else
126         lblWAPPushSMSPictureDetails.caption = ""
128         picContainer.Visible = False
        End If

130     Form_Resize
        '<EhFooter>
        Exit Sub

Form_Paint_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.Form_Paint " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim lTemp As Long
    Dim rColWidthTotal As Long
    Dim rTemp As Double
    Dim i As Integer
    Dim bWAPPushHorizontalScrollBarRequired As Boolean
    Dim bWAPPushVerticalScrollBarRequired As Boolean
    Dim lWAPPushContainerPossibleMaxWidth As Long
    Dim lWAPPushContainerPossibleMaxHeight As Long
    Dim lTabMainWidthCache As Long

    If Me.WindowState = 1 Then Exit Sub

    VersionSpecificAction 59

    lTemp = Me.ScaleWidth - (TabMain.Left * 2)
    lTabMainWidthCache = IIf(lTemp > 0, lTemp, 0)
    TabMain.Width = lTabMainWidthCache
  
    lTemp = Me.ScaleHeight - TabMain.Left
    TabMain.Height = IIf(lTemp > 0, lTemp, 0)

    Select Case TabMain.Tab

        Case 0  'General
            lTemp = (TabMain.Width - (fraSMSType.Left * 3)) / 2
            fraSMSType.Width = IIf(lTemp > 0, lTemp, 0)
  
            For i = optSMSType.LBound To optSMSType.UBound
                optSMSType(i).Width = fraSMSType.Width - optSMSType(i).Left - 50
            Next
  
            lTemp = fraSMSType.Width
            fraOptions.Width = IIf(lTemp > 0, lTemp, 0)
  
'            lTemp = fraOptions.Width - (cmdDeferredDeliveryTime.Left * 2)
'            cmdDeferredDeliveryTime.Width = IIf(lTemp > 0, lTemp, 0)
'
'            lTemp = fraOptions.Width - (cmdSend.Left * 2)
'            cmdSend.Width = IIf(lTemp > 0, lTemp, 0)

            lTemp = fraSMSType.Width
            fraCurrentJob.Width = IIf(lTemp > 0, lTemp, 0)
  
            lTemp = TabMain.Height - fraCurrentJob.Top - fraSMSType.Left
            fraCurrentJob.Height = IIf(lTemp > 0, lTemp, 0)
            fraCurrentJob.Height = fraCurrentJob.Height - (pctLogo.Height + 80)
  
            lTemp = fraSMSType.Left + fraSMSType.Width + fraSMSType.Left
            fraCurrentJob.Left = IIf(lTemp > 0, lTemp, 0)

            lTemp = fraCurrentJob.Width - (lstCurrentJob.Left * 2)
            lstCurrentJob.Width = IIf(lTemp > 0, lTemp, 0)
  
            lTemp = fraCurrentJob.Height - (lstCurrentJob.Top * 1.1)
            lstCurrentJob.Height = IIf(lTemp > 0, lTemp, 0)

        Case 1  'Recipients
            lTemp = (TabMain.Width - (fraCurrentRecipients.Left * 3)) / 2
            fraCurrentRecipients.Width = IIf(lTemp > 0, lTemp, 0)
  
            lTemp = TabMain.Height - fraCurrentRecipients.Top - fraCurrentRecipients.Left
            fraCurrentRecipients.Height = IIf(lTemp > 0, lTemp, 0)

            lTemp = fraCurrentRecipients.Width
            fraPhonebook.Width = IIf(lTemp > 0, lTemp, 0)
  
            lTemp = TabMain.Height - fraPhonebook.Top - fraCurrentRecipients.Left
            fraPhonebook.Height = IIf(lTemp > 0, lTemp, 0)
  
            lTemp = fraCurrentRecipients.Left + fraCurrentRecipients.Width + fraCurrentRecipients.Left
            fraPhonebook.Left = IIf(lTemp > 0, lTemp, 0)

            ResizerResizeControl cmdAddToRecipientList, fraCurrentRecipients, 1
'            ResizerResizeControl cmdImportRecipientsFromFile, fraCurrentRecipients, 1
            ResizerResizeControl cmdRemove, fraCurrentRecipients, 1
           ResizerResizeControl cmdClearList, fraCurrentRecipients, 1
            ResizerResizeControl cmdDuplicates, fraCurrentRecipients, 1
  
            lTemp = cmdAddToRecipientList.Left - (grdRecipients.Left * 2)
            grdRecipients.Width = IIf(lTemp > 0, lTemp, 0)

            lTemp = fraCurrentRecipients.Height - grdRecipients.Top - grdRecipients.Left
            grdRecipients.Height = IIf(lTemp > 0, lTemp, 0)
  
            lTemp = grdRecipients.Width
            txtRecipient.Width = IIf(lTemp > 0, lTemp, 0)
  
            lTemp = fraPhonebook.Width - grdPhonebook.Left - cmdAddNew.Width
            cmdAddNew.Left = IIf(lTemp > 0, lTemp, 0)
            cmdUpdate.Left = IIf(lTemp > 0, lTemp, 0)
            cmdDelete.Left = IIf(lTemp > 0, lTemp, 0)
  
            lTemp = fraPhonebook.Width - (grdPhonebook.Left * 2)
            grdPhonebook.Width = IIf(lTemp > 0, lTemp, 0)
  
            lTemp = fraPhonebook.Height - grdPhonebook.Top - grdPhonebook.Left
            grdPhonebook.Height = IIf(lTemp > 0, lTemp, 0)
  
            ResizerResizeControl txtPhonebookName, cmdAddNew, 2
            ResizerResizeControl txtPhonebookPhoneNumber, cmdAddNew, 2

            For i = txtPhonebookVariableField.LBound To txtPhonebookVariableField.UBound
                ResizerResizeControl txtPhonebookVariableField(i), cmdAddNew, 2
            Next

            ResizerResizeControl txtFilter, cmdAddNew, 2

            lTemp = txtFilter.Width + txtFilter.Left - cmdAddRecipientsFromPhonebook.Left
            cmdAddRecipientsFromPhonebook.Width = IIf(lTemp > 0, lTemp, 0)
  
            rColWidthTotal = 0

            For i = 0 To grdRecipients.Cols - 1
                rColWidthTotal = rColWidthTotal + grdRecipients.ColWidth(i)
            Next
  
            For i = 0 To grdRecipients.Cols - 1
                grdRecipients.ColWidth(i) = grdRecipients.ColWidth(i) / rColWidthTotal * (grdRecipients.Width - 380)
            Next
  
            rColWidthTotal = 0

            For i = 0 To grdPhonebook.Cols - 1
                rColWidthTotal = rColWidthTotal + grdPhonebook.ColWidth(i)
            Next
  
            For i = 0 To grdPhonebook.Cols - 1
                grdPhonebook.ColWidth(i) = grdPhonebook.ColWidth(i) / rColWidthTotal * (grdPhonebook.Width - 380)
            Next
  
        Case 10  'WAP Push Picture
            lWAPPushContainerPossibleMaxWidth = lTabMainWidthCache - picContainer.Left - picContainer.Left
            lWAPPushContainerPossibleMaxHeight = TabMain.Height - picContainer.Top - picContainer.Left
            imgPreviewWAPPushSMSPicture.Left = 0
            imgPreviewWAPPushSMSPicture.Top = 0
  
            If imgPreviewWAPPushSMSPicture.Width > lWAPPushContainerPossibleMaxWidth Then
                bWAPPushHorizontalScrollBarRequired = True
            Else
                bWAPPushHorizontalScrollBarRequired = False
            End If

            If imgPreviewWAPPushSMSPicture.Height > lWAPPushContainerPossibleMaxHeight Then
                bWAPPushVerticalScrollBarRequired = True
            Else
                bWAPPushVerticalScrollBarRequired = False
            End If
  
            If bWAPPushHorizontalScrollBarRequired = True Then
                picContainer.Width = lTabMainWidthCache - picContainer.Left - picContainer.Left
                scrlHorizontal.Min = 0
                lTemp = imgPreviewWAPPushSMSPicture.Width - picContainer.ScaleWidth
                scrlHorizontal.Max = IIf(lTemp > 32000, 32000, lTemp)
                scrlHorizontal.SmallChange = 10
                lTemp = imgPreviewWAPPushSMSPicture.Width + picContainer.ScaleWidth
                scrlHorizontal.LargeChange = IIf(lTemp > 32000, 32000, lTemp)
                scrlHorizontal.Value = 0
            Else
                picContainer.Width = imgPreviewWAPPushSMSPicture.Width
            End If
  
            If bWAPPushVerticalScrollBarRequired = True Then
                lTemp = TabMain.Height - picContainer.Top - picContainer.Left
                picContainer.Height = IIf(lTemp < 0, 0, lTemp)
                scrlVertical.Min = 0
                lTemp = imgPreviewWAPPushSMSPicture.Height - picContainer.ScaleHeight
                scrlVertical.Max = IIf(lTemp > 32000, 32000, lTemp)
                lTemp = imgPreviewWAPPushSMSPicture.Height + picContainer.ScaleHeight
                scrlVertical.LargeChange = IIf(lTemp > 32000, 32000, lTemp)
                lTemp = scrlVertical.LargeChange / 50
                scrlVertical.SmallChange = IIf(lTemp < 1, 1, lTemp)
                scrlVertical.Value = 0
            Else
                picContainer.Height = imgPreviewWAPPushSMSPicture.Height
            End If
  
            scrlHorizontal.Left = 0
            scrlHorizontal.Height = 255
            scrlVertical.Top = 0
            scrlVertical.Width = 255
  
            Select Case True

                Case bWAPPushHorizontalScrollBarRequired = False And bWAPPushVerticalScrollBarRequired = False
                    scrlHorizontal.Visible = False
                    scrlVertical.Visible = False
                    picHideBottomRightCorner.Visible = False
  
                Case bWAPPushHorizontalScrollBarRequired And bWAPPushVerticalScrollBarRequired
                    scrlHorizontal.Width = picContainer.ScaleWidth - scrlHorizontal.Height
                    scrlHorizontal.Top = picContainer.ScaleHeight - scrlHorizontal.Height
                    scrlVertical.Left = picContainer.ScaleWidth - scrlVertical.Width
                    lTemp = picContainer.ScaleHeight - scrlVertical.Width
                    scrlVertical.Height = IIf(lTemp < 0, 0, lTemp)
                    scrlHorizontal.Visible = True
                    scrlVertical.Visible = True
                    picHideBottomRightCorner.Top = scrlVertical.Height
                    picHideBottomRightCorner.Left = scrlHorizontal.Width
                    picHideBottomRightCorner.BorderStyle = vbBSNone
                    picHideBottomRightCorner.Visible = True
    
                Case bWAPPushHorizontalScrollBarRequired
                    scrlHorizontal.Width = picContainer.ScaleWidth
                    scrlHorizontal.Top = picContainer.ScaleHeight - scrlHorizontal.Height
                    scrlHorizontal.Visible = True
                    scrlVertical.Visible = False
                    picHideBottomRightCorner.Visible = False
    
                Case bWAPPushVerticalScrollBarRequired
                    scrlVertical.Left = picContainer.ScaleWidth - scrlVertical.Width
                    scrlVertical.Height = picContainer.ScaleHeight
                    scrlHorizontal.Visible = False
                    scrlVertical.Visible = True
                    picHideBottomRightCorner.Visible = False
    
            End Select
    End Select

End Sub

Private Sub grdPhonebook_EnterCell()
        '<EhHeader>
        On Error GoTo grdPhonebook_EnterCell_Err
        '</EhHeader>
100     grdPhonebook_Click
        '<EhFooter>
        Exit Sub

grdPhonebook_EnterCell_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.grdPhonebook_EnterCell " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub grdPhonebook_KeyDown(KeyCode As Integer, _
                                 Shift As Integer)
        '<EhHeader>
        On Error GoTo grdPhonebook_KeyDown_Err
        '</EhHeader>

100     Select Case True

            Case KeyCode = 67 And Shift = 2  'CTRL C
102             CopySelectedGridAreaToClipboard grdPhonebook
  
104         Case KeyCode = 45 And Shift = 2  'CTRL INS
106             CopySelectedGridAreaToClipboard grdPhonebook
  
108         Case Else
                'Do nothing
  
        End Select

        '<EhFooter>
        Exit Sub

grdPhonebook_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.grdPhonebook_KeyDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub grdPhonebook_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, _
                                   y As Single)
        '<EhHeader>
        On Error GoTo grdPhonebook_MouseDown_Err
        '</EhHeader>

100     If Button = vbRightButton Then
102         Me.PopupMenu Me.mnuEditTop01
        End If

        Exit Sub

104     If grdPhonebook.Row <= 1 And Button = 2 Then
   
106         Select Case True

                Case grdPhonebook.col = 1 And gnPhoneBookSortOrder = 1 'Switch from Name ASC to Name DESC
108                 gnPhoneBookSortOrder = 2
    
110             Case grdPhonebook.col = 1 And gnPhoneBookSortOrder = 2 'Switch from Name DESC to Name ASC
112                 gnPhoneBookSortOrder = 1
    
114             Case grdPhonebook.col = 2 And gnPhoneBookSortOrder = 3 'Switch from Number ASC to Number DESC
116                 gnPhoneBookSortOrder = 4
    
118             Case grdPhonebook.col = 2 And gnPhoneBookSortOrder = 4 'Switch from Number DESC to Number ASC
120                 gnPhoneBookSortOrder = 3
    
122             Case grdPhonebook.col = 1
124                 gnPhoneBookSortOrder = 1
    
126             Case grdPhonebook.col = 2
128                 gnPhoneBookSortOrder = 3
  
            End Select
  
130         txtFilter_Change
        Else

        End If

        '<EhFooter>
        Exit Sub

grdPhonebook_MouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.grdPhonebook_MouseDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub grdRecipients_DblClick()
        '<EhHeader>
        On Error GoTo grdRecipients_DblClick_Err
        '</EhHeader>
100     cmdRemove_Click
        '<EhFooter>
        Exit Sub

grdRecipients_DblClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.grdRecipients_DblClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub grdRecipients_KeyDown(KeyCode As Integer, _
                                  Shift As Integer)
        '<EhHeader>
        On Error GoTo grdRecipients_KeyDown_Err
        '</EhHeader>

100     Select Case True

            Case KeyCode = 67 And Shift = 2  'CTRL C
102             CopySelectedGridAreaToClipboard grdRecipients
  
104         Case KeyCode = 45 And Shift = 2  'CTRL INS
106             CopySelectedGridAreaToClipboard grdRecipients
  
108         Case Else
                'Do nothing
  
        End Select

        '<EhFooter>
        Exit Sub

grdRecipients_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.grdRecipients_KeyDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub grdRecipients_MouseDown(Button As Integer, _
                                    Shift As Integer, _
                                    x As Single, _
                                    y As Single)
        '<EhHeader>
        On Error GoTo grdRecipients_MouseDown_Err
        '</EhHeader>

100     If Button = vbRightButton Then
102         Me.PopupMenu Me.mnuEditTop02
        End If

        '<EhFooter>
        Exit Sub

grdRecipients_MouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.grdRecipients_MouseDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub grdTemplates_Click()
        '<EhHeader>
        On Error GoTo grdTemplates_Click_Err
        '</EhHeader>

100     If grdTemplates.Rows <= 1 Then Exit Sub

102     mlCurrentTemplateID = grdTemplates.TextMatrix(grdTemplates.RowSel, 0)

104     dtaTemplates.Recordset.FindFirst "ID = " & mlCurrentTemplateID

106     txtTemplateMessage.Text = dtaTemplates.Recordset("Message") & ""
        '<EhFooter>
        Exit Sub

grdTemplates_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.grdTemplates_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub grdTemplates_EnterCell()
        '<EhHeader>
        On Error GoTo grdTemplates_EnterCell_Err
        '</EhHeader>
100     grdTemplates_Click
        '<EhFooter>
        Exit Sub

grdTemplates_EnterCell_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.grdTemplates_EnterCell " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub imgRandomLogo_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo imgRandomLogo_Click_Err
        '</EhHeader>
        Dim sFilename As String
        Dim i As Integer

        On Error GoTo ErrorTrap
100     sFilename = App.Path & "\" & Right("0000" & Trim(Str$(Index)), 4) & ".bmp"
102     imgSelectedRandomLogo.Picture = LoadPicture(sFilename)

104     gnSelectedRandomLogo = Index

        Exit Sub
ErrorTrap:
        Exit Sub
        '<EhFooter>
        Exit Sub

imgRandomLogo_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.imgRandomLogo_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub lstCountrys_Click()
        '<EhHeader>
        On Error GoTo lstCountrys_Click_Err
        '</EhHeader>
100     lstOperators.Clear

        'Search for all networks for one specific country, first match is a special case and therefore handled separately
102     dtaNetworks.Recordset.FindFirst "Country = '" & lstCountrys.List(lstCountrys.ListIndex) & "'"

104     If dtaNetworks.Recordset.NoMatch Then
            'Do nothing
        Else
106         lstOperators.AddItem dtaNetworks.Recordset("Operator") & ""
108         lstOperators.ItemData(lstOperators.NewIndex) = dtaNetworks.Recordset("ID")
        End If

        'Search for all other networks
110     Do While Not dtaNetworks.Recordset.EOF
112         dtaNetworks.Recordset.FindNext "Country = '" & lstCountrys.List(lstCountrys.ListIndex) & "'"

114         If dtaNetworks.Recordset.NoMatch Then
                Exit Do
            Else
116             lstOperators.AddItem dtaNetworks.Recordset("Operator") & ""
118             lstOperators.ItemData(lstOperators.NewIndex) = dtaNetworks.Recordset("ID")
            End If

        Loop

120     If lstOperators.ListCount >= 1 Then
122         lstOperators.ListIndex = 0
        End If

        '<EhFooter>
        Exit Sub

lstCountrys_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.lstCountrys_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub lstOperators_Click()
        '<EhHeader>
        On Error GoTo lstOperators_Click_Err
        '</EhHeader>
100     dtaNetworks.Recordset.FindFirst "ID = " & lstOperators.ItemData(lstOperators.ListIndex)
102     txtMCC.Text = dtaNetworks.Recordset("MCC") & ""
104     txtMNC.Text = dtaNetworks.Recordset("MNC") & ""
        '<EhFooter>
        Exit Sub

lstOperators_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.lstOperators_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuInfo009_Click()
        '<EhHeader>
        On Error GoTo mnuInfo009_Click_Err
        '</EhHeader>
100     VersionSpecificAction 35
        '<EhFooter>
        Exit Sub

mnuInfo009_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuInfo009_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuAddRecipientsFromPhonebook01_Click()
        '<EhHeader>
        On Error GoTo mnuAddRecipientsFromPhonebook01_Click_Err
        '</EhHeader>
100     cmdAddRecipientsFromPhonebook_Click
        '<EhFooter>
        Exit Sub

mnuAddRecipientsFromPhonebook01_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuAddRecipientsFromPhonebook01_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuInfo002_Click()
        '<EhHeader>
        On Error GoTo mnuInfo002_Click_Err
        '</EhHeader>
100     VersionSpecificAction 28
        '<EhFooter>
        Exit Sub

mnuInfo002_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuInfo002_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuBuyCredits_Click()
        '<EhHeader>
        On Error GoTo mnuBuyCredits_Click_Err
        '</EhHeader>
100     VersionSpecificAction 54
        '<EhFooter>
        Exit Sub

mnuBuyCredits_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuBuyCredits_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuClearList02_Click()
        '<EhHeader>
        On Error GoTo mnuClearList02_Click_Err
        '</EhHeader>
100     cmdClearList_Click
        '<EhFooter>
        Exit Sub

mnuClearList02_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuClearList02_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuCopy01_Click()
        '<EhHeader>
        On Error GoTo mnuCopy01_Click_Err
        '</EhHeader>
100     CopySelectedGridAreaToClipboard grdPhonebook
        '<EhFooter>
        Exit Sub

mnuCopy01_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuCopy01_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuCopy02_Click()
        '<EhHeader>
        On Error GoTo mnuCopy02_Click_Err
        '</EhHeader>
100     CopySelectedGridAreaToClipboard grdRecipients
        '<EhFooter>
        Exit Sub

mnuCopy02_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuCopy02_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuDelete_Click()
        '<EhHeader>
        On Error GoTo mnuDelete_Click_Err
        '</EhHeader>
100     cmdDelete_Click
        '<EhFooter>
        Exit Sub

mnuDelete_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuDelete_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuDuplicates02_Click()
        '<EhHeader>
        On Error GoTo mnuDuplicates02_Click_Err
        '</EhHeader>
100     cmdDuplicates_Click
        '<EhFooter>
        Exit Sub

mnuDuplicates02_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuDuplicates02_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuExit_Click()
        '<EhHeader>
        On Error GoTo mnuExit_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

mnuExit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuExit_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuInfo007_Click()
        '<EhHeader>
        On Error GoTo mnuInfo007_Click_Err
        '</EhHeader>
100     VersionSpecificAction 33
        '<EhFooter>
        Exit Sub

mnuInfo007_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuInfo007_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuJoblog_Click()
'        '<EhHeader>
'        On Error GoTo mnuJoblog_Click_Err
'        '</EhHeader>
'100     Unload frmJoblog
'102     frmJoblog.Show vbModal, Me
'        '<EhFooter>
'        Exit Sub
'
'mnuJoblog_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmSMSMain.mnuJoblog_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub mnuLanguage_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo mnuLanguage_Click_Err
        '</EhHeader>
        Dim frmTemp As Form

100     Select Case Index

            Case 1
102             mnuLanguage(1).Checked = True
104             mnuLanguage(2).Checked = False
106             gnLanguage = 1
  
108         Case 2
110             mnuLanguage(1).Checked = False
112             mnuLanguage(2).Checked = True
114             gnLanguage = 2
  
        End Select

116     For Each frmTemp In Forms
118         frmTemp.AdjustLanguageSettings gnLanguage
        Next

        '<EhFooter>
        Exit Sub

mnuLanguage_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuLanguage_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuInfo004_Click()
        '<EhHeader>
        On Error GoTo mnuInfo004_Click_Err
        '</EhHeader>
100     VersionSpecificAction 30
        '<EhFooter>
        Exit Sub

mnuInfo004_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuInfo004_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuMarkAll01_Click()
        '<EhHeader>
        On Error GoTo mnuMarkAll01_Click_Err
        '</EhHeader>

100     If grdPhonebook.Rows > 1 Then
102         grdPhonebook.Row = 1
104         grdPhonebook.col = 0
106         grdPhonebook.RowSel = grdPhonebook.Rows - 1
108         grdPhonebook.ColSel = grdPhonebook.Cols - 1
        End If

        '<EhFooter>
        Exit Sub

mnuMarkAll01_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuMarkAll01_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuMarkAll02_Click()
        '<EhHeader>
        On Error GoTo mnuMarkAll02_Click_Err
        '</EhHeader>

100     If grdRecipients.Rows > 1 Then
102         grdRecipients.Row = 1
104         grdRecipients.col = 0
106         grdRecipients.RowSel = grdRecipients.Rows - 1
108         grdRecipients.ColSel = grdRecipients.Cols - 1
        End If

        '<EhFooter>
        Exit Sub

mnuMarkAll02_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuMarkAll02_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuInfo003_Click()
        '<EhHeader>
        On Error GoTo mnuInfo003_Click_Err
        '</EhHeader>
100     VersionSpecificAction 29
        '<EhFooter>
        Exit Sub

mnuInfo003_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuInfo003_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuInfo006_Click()
        '<EhHeader>
        On Error GoTo mnuInfo006_Click_Err
        '</EhHeader>
100     VersionSpecificAction 32
        '<EhFooter>
        Exit Sub

mnuInfo006_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuInfo006_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuOpenDatabase_Click()
        '<EhHeader>
        On Error GoTo mnuOpenDatabase_Click_Err
        '</EhHeader>
        Dim sTemp As String
        Dim sApplicationPlaceHolder As String
        On Error GoTo ErrorCancelSelected

100     VersionSpecificAction 44, , , sApplicationPlaceHolder
102     sTemp = LoadLanguageSpecificString(gnLanguage, 342)
104     sTemp = Replace(sTemp, gcPlaceHolder, sApplicationPlaceHolder)

106     cmdDlgOpen.Filter = sTemp
108     cmdDlgOpen.ShowOpen

110     gsPathAndDatabaseName = cmdDlgOpen.Filename

112     InitApp

        Exit Sub
ErrorCancelSelected:
        Exit Sub

        '<EhFooter>
        Exit Sub

mnuOpenDatabase_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuOpenDatabase_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuOpenDatabaseWithAccess2000_Click()
        '<EhHeader>
        On Error GoTo mnuOpenDatabaseWithAccess2000_Click_Err
        '</EhHeader>
        Dim lRet As Long
100     lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", gsPathAndDatabaseName, vbNullString, vbNullString, vbNormalFocus)
        '<EhFooter>
        Exit Sub

mnuOpenDatabaseWithAccess2000_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuOpenDatabaseWithAccess2000_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuOriginators_Click()
'        '<EhHeader>
'        On Error GoTo mnuOriginators_Click_Err
'        '</EhHeader>
'100     frmOriginatorCheck.txtStep1Originator = txtOriginator.Text
'102     frmOriginatorCheck.Show vbModal, Me
'        '<EhFooter>
'        Exit Sub
'
'mnuOriginators_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmSMSMain.mnuOriginators_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub mnuRegistration_Click()
'        '<EhHeader>
'        On Error GoTo mnuRegistration_Click_Err
'        '</EhHeader>
'100     VersionSpecificAction 55
'        '<EhFooter>
'        Exit Sub
'
'mnuRegistration_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmSMSMain.mnuRegistration_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub mnuRemoveFromRecipientList02_Click()
        '<EhHeader>
        On Error GoTo mnuRemoveFromRecipientList02_Click_Err
        '</EhHeader>
100     cmdRemove_Click
        '<EhFooter>
        Exit Sub

mnuRemoveFromRecipientList02_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuRemoveFromRecipientList02_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuSendlog_Click()
'        '<EhHeader>
'        On Error GoTo mnuSendlog_Click_Err
'        '</EhHeader>
'100     gtSendLogFilterSettings.bJobIDFilterUsed = False
'102     gtSendLogFilterSettings.lJobId = 0
'104     Unload frmSendLog
'106     frmSendLog.Show vbModal, Me
'        '<EhFooter>
'        Exit Sub
'
'mnuSendlog_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmSMSMain.mnuSendlog_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub mnuShowCredits_Click()
'        '<EhHeader>
'        On Error GoTo mnuShowCredits_Click_Err
'        '</EhHeader>
'100     ShowCredits True
'        '<EhFooter>
'        Exit Sub
'
'mnuShowCredits_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmSMSMain.mnuShowCredits_Click " & _
'               "at line " & Erl
'        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuInfo001_Click()
        '<EhHeader>
        On Error GoTo mnuInfo001_Click_Err
        '</EhHeader>
100     VersionSpecificAction 56
        '<EhFooter>
        Exit Sub

mnuInfo001_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuInfo001_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuInfo008_Click()
        '<EhHeader>
        On Error GoTo mnuInfo008_Click_Err
        '</EhHeader>
100     VersionSpecificAction 34
        '<EhFooter>
        Exit Sub

mnuInfo008_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuInfo008_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuInfo005_Click()
        '<EhHeader>
        On Error GoTo mnuInfo005_Click_Err
        '</EhHeader>
100     VersionSpecificAction 31
        '<EhFooter>
        Exit Sub

mnuInfo005_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.mnuInfo005_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub optSMSType_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo optSMSType_Click_Err
        '</EhHeader>
        Dim nEnabledTab As Integer
        Dim i As Integer

100     Select Case Index
 
            Case 0 To 11
102             cmdSend.Enabled = True
104             cmdSend.caption = LoadLanguageSpecificString(gnLanguage, 171) & " " & SelectedSMSType(gnLanguage)
106             UpdateJobList gnLanguage
  
108         Case Else
                'Unexecpted, do nothing, probabely an application bug
  
        End Select

110     Select Case Index

            Case 0
112             nEnabledTab = 2
  
114         Case 1
116             nEnabledTab = 3
  
118         Case 2
120             nEnabledTab = 3
  
122         Case 3
124             nEnabledTab = 4
  
126         Case 4
128             nEnabledTab = 5
  
130         Case 5
132             nEnabledTab = 6
  
134         Case 6
136             nEnabledTab = 9
  
138         Case 7
140             nEnabledTab = 10
  
142         Case 8
144             nEnabledTab = 7
  
146         Case 9
148             nEnabledTab = 8
  
150         Case 10
152             nEnabledTab = -1
  
154         Case 11
156             nEnabledTab = -1
  
158         Case Else
                'MsgBox "Case Else"
  
        End Select

160     For i = 2 To 10

162         If i = nEnabledTab Then
164             frmSMSMain.TabMain.TabVisible(i) = True
            Else
166             frmSMSMain.TabMain.TabVisible(i) = False
            End If

        Next

168     Select Case nEnabledTab

            Case 3, 4, 5, 10
170             frmSMSMain.TabMain.OLEDropMode = ssOLEDropManual
172             frmSMSMain.OLEDropMode = ssOLEDropManual
  
174         Case Else
176             frmSMSMain.TabMain.OLEDropMode = ssOLEDropManual
178             frmSMSMain.OLEDropMode = ssOLEDropManual
  
        End Select
  
        '<EhFooter>
        Exit Sub

optSMSType_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.optSMSType_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAddNew_Click()
        '<EhHeader>
        On Error GoTo cmdAddNew_Click_Err
        '</EhHeader>
        Dim lIDInsertedRow As Long
        Dim lCurrentRow As Long
        Dim lMaxRows As Long
        Dim lHeightOfAllRowsAbove As Long

100     dtaPhonebook.Recordset.AddNew
102     dtaPhonebook.Recordset(LoadLanguageSpecificString(gnLanguage, 196)) = txtPhonebookName.Text
104     dtaPhonebook.Recordset(LoadLanguageSpecificString(gnLanguage, 197)) = txtPhonebookPhoneNumber.Text
106     dtaPhonebook.Recordset(gsPhonebookVariableField(gnLanguage, 1)) = txtPhonebookVariableField(1).Text
108     dtaPhonebook.Recordset(gsPhonebookVariableField(gnLanguage, 2)) = txtPhonebookVariableField(2).Text
110     dtaPhonebook.Recordset(gsPhonebookVariableField(gnLanguage, 3)) = txtPhonebookVariableField(3).Text
112     dtaPhonebook.Recordset.UpDate

114     dtaPhonebook.Recordset.MoveLast
116     lIDInsertedRow = dtaPhonebook.Recordset("ID")
118     dtaPhonebook.Refresh

120     DoEvents
122     grdPhonebook.Refresh

124     SetCorrectTopRowWithinPhonebook lIDInsertedRow

126     grdPhonebook.Refresh

128     mlCurrentPhoneBookID = lIDInsertedRow
        '<EhFooter>
        Exit Sub

cmdAddNew_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdAddNew_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdClearList_Click()
        '<EhHeader>
        On Error GoTo cmdClearList_Click_Err
        '</EhHeader>
        Dim i As Long
100     grdRecipients.Rows = 1
102     UpdateJobList gnLanguage
        '<EhFooter>
        Exit Sub

cmdClearList_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdClearList_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdDelete_Click()
        '<EhHeader>
        On Error GoTo cmdDelete_Click_Err
        '</EhHeader>
        Dim lStart As Long
        Dim lEnd As Long
        Dim nAnswer As Integer
        Dim i As Long

        On Error GoTo ErrorTrap

100     If grdPhonebook.Rows < 2 Then Exit Sub

102     If grdPhonebook.RowSel < grdPhonebook.Row Then
104         lStart = grdPhonebook.RowSel
106         lEnd = grdPhonebook.Row
        Else
108         lStart = grdPhonebook.Row
110         lEnd = grdPhonebook.RowSel
        End If

112     nAnswer = MsgBox(LoadLanguageSpecificString(gnLanguage, 208) & Str$(lEnd - lStart + 1) & " " & LoadLanguageSpecificString(gnLanguage, 209), vbOKCancel Or vbQuestion, gsApplicationName)
          
114     If nAnswer = vbOK Then
116         Screen.MousePointer = vbHourglass

118         For i = lStart To lEnd
120             dtaPhonebook.Recordset.FindFirst "ID = " & grdPhonebook.TextMatrix(i, 0)
122             dtaPhonebook.Recordset.Delete
            Next

124         dtaPhonebook.Refresh
126         Screen.MousePointer = vbDefault
        End If

128     txtFilter_Change

ErrorTrap:
130     Screen.MousePointer = vbDefault
        Exit Sub
        '<EhFooter>
        Exit Sub

cmdDelete_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdDelete_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdRemove_Click()
        '<EhHeader>
        On Error GoTo cmdRemove_Click_Err
        '</EhHeader>
        Dim lStart As Long
        Dim lEnd As Long
        Dim i As Long

100     If grdRecipients.Rows <= 1 Then
            Exit Sub
        End If

102     If grdRecipients.RowSel < grdRecipients.Row Then
104         lEnd = grdRecipients.RowSel
106         lStart = grdRecipients.Row
        Else
108         lEnd = grdRecipients.Row
110         lStart = grdRecipients.RowSel
        End If

112     If lStart - lEnd + 1 = grdRecipients.Rows - 1 Then
114         grdRecipients.Rows = 1
        Else

116         For i = lStart To lEnd Step -1
118             grdRecipients.RemoveItem i
            Next

        End If

120     grdRecipients.RowSel = grdRecipients.Row
122     UpdateJobList gnLanguage
        '<EhFooter>
        Exit Sub

cmdRemove_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdRemove_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSelectHandylogo_Click()
        '<EhHeader>
        On Error GoTo cmdSelectHandylogo_Click_Err
        '</EhHeader>
        On Error GoTo ErrorCancelSelected

100     cmdDlgOpen.Filter = LoadLanguageSpecificString(gnLanguage, 201)
102     cmdDlgOpen.ShowOpen
104     txtPathLogo.Text = cmdDlgOpen.Filename
106     imgPreviewOperatorLogo.Picture = LoadPicture(cmdDlgOpen.Filename)

        Exit Sub
ErrorCancelSelected:
        Exit Sub
        '<EhFooter>
        Exit Sub

cmdSelectHandylogo_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdSelectHandylogo_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSelectPictureMessage_Click()
        '<EhHeader>
        On Error GoTo cmdSelectPictureMessage_Click_Err
        '</EhHeader>
        On Error GoTo ErrorCancelSelected

100     cmdDlgOpen.Filter = LoadLanguageSpecificString(gnLanguage, 201)
102     cmdDlgOpen.ShowOpen
104     txtPathPictureMessage.Text = cmdDlgOpen.Filename
106     imgPreviewPictureMessage.Picture = LoadPicture(cmdDlgOpen.Filename)

        Exit Sub
ErrorCancelSelected:
        Exit Sub
        '<EhFooter>
        Exit Sub

cmdSelectPictureMessage_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdSelectPictureMessage_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSelectRingtone_Click()
        '<EhHeader>
        On Error GoTo cmdSelectRingtone_Click_Err
        '</EhHeader>
        On Error GoTo ErrorCancelSelected

100     cmdDlgOpen.Filter = LoadLanguageSpecificString(gnLanguage, 202)
102     cmdDlgOpen.ShowOpen
104     txtPathRingtone = cmdDlgOpen.Filename

        Exit Sub
ErrorCancelSelected:
        Exit Sub
        '<EhFooter>
        Exit Sub

cmdSelectRingtone_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdSelectRingtone_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSend_Click()
        '<EhHeader>
        On Error GoTo cmdSend_Click_Err
        '</EhHeader>
        Dim sMessage As String
        Dim nAnswer As Integer
        Dim bSuccess As Boolean
        Dim dWork As Date
        Dim k As Long

        Dim lRecipient As Long
        Dim lRecipientStart As Long
        Dim lRecipientEnd As Long
        Dim nStepSize As Integer
        Dim lCurrentJobRecipient As Long
        Dim bMultipleStepJob As Boolean
        Dim lJobId As Long
        Dim lDuplicates As Long
        Dim bJobSuccessfullyProcessed As Boolean
        Dim cRecipientData As New Collection
        Dim cRecipientTemp As New Recipient
        Dim lUpperLimit As Long
        Dim lStepCounter As Long
        Dim lTotalSteps As Long
        Dim lPeriodicSendDate As Long

        Dim bVariableMessageData As Boolean
        Dim bVariableDeferredDeliveryTimeData As Boolean

        Dim sRecipient As String
        Dim sRecipientNameFromGrid As String
        Dim sTransactionReferenceNumber As String
        Dim sMessageData As String

100     If grdRecipients.Rows <= 1 Then
102         sMessage = LoadLanguageSpecificString(gnLanguage, 489)
104         MsgBox sMessage, vbInformation, gsApplicationName
            Exit Sub
        End If

106     If DuplicatesFoundWithinCurrentRecipients(lDuplicates) = True Then
            'Warningmessage DUPLICATES FOUND
108         sMessage = LoadLanguageSpecificString(gnLanguage, 491) & " " & Trim(Str$(lDuplicates)) & " " & LoadLanguageSpecificString(gnLanguage, 492) & vbCrLf & LoadLanguageSpecificString(gnLanguage, 493)
110         nAnswer = MsgBox(sMessage, vbExclamation Or vbOKCancel Or vbDefaultButton2, gsApplicationName)
  
112         If nAnswer = vbCancel Then
                Exit Sub
            End If
        End If

114     If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True Then
116         If gtDeferredDeliveryTimeSettings.bSingleSMS = False Then
118             If IsDate(gtDeferredDeliveryTimeSettings.sStartingDate) Then
120                 dWork = CVDate(gtDeferredDeliveryTimeSettings.sStartingDate)
                End If

            Else

122             If IsDate(gtDeferredDeliveryTimeSettings.sDeliveryDate) Then
124                 dWork = CVDate(gtDeferredDeliveryTimeSettings.sDeliveryDate)
                End If
            End If

126         If dWork < Now() Then
                'Warningmessage DEFERRED DELIVERY TIME IN THE PAST
128             sMessage = LoadLanguageSpecificString(gnLanguage, 494) & vbCrLf & LoadLanguageSpecificString(gnLanguage, 495)
130             nAnswer = MsgBox(sMessage, vbExclamation Or vbOKCancel Or vbDefaultButton2, gsApplicationName)
  
132             If nAnswer = vbCancel Then
                    Exit Sub
                End If
            End If
        End If

134     If CheckOriginatorBeforeSending() = False Then
            Exit Sub
        End If

136     sMessage = LoadLanguageSpecificString(gnLanguage, 210) & SelectedSMSType(gnLanguage) & LoadLanguageSpecificString(gnLanguage, 211)
138     nAnswer = MsgBox(sMessage, vbOKCancel Or vbInformation, gsApplicationName)

140     If nAnswer <> vbOK Then
            Exit Sub
        End If

142     If gbSaveMessagesInSendLog = True Then
144         lJobId = SaveCurrentJobInfo(gnLanguage)
        End If

146     nStepSize = 50

148     If grdRecipients.Rows - 1 > nStepSize Then
150         bMultipleStepJob = True
        End If

152     bJobSuccessfullyProcessed = True

154     For lRecipient = 1 To grdRecipients.Rows - 1
       
156         If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True And gtDeferredDeliveryTimeSettings.bSingleSMS = False Then
158             lUpperLimit = Val(gtDeferredDeliveryTimeSettings.sNumberOfMessages)
160             bVariableDeferredDeliveryTimeData = True

162             If IsDate(gtDeferredDeliveryTimeSettings.sStartingDate) Then
164                 dWork = CVDate(gtDeferredDeliveryTimeSettings.sStartingDate)
                End If

            Else
166             lUpperLimit = 1
168             bVariableDeferredDeliveryTimeData = False

170             If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True Then
172                 If IsDate(gtDeferredDeliveryTimeSettings.sDeliveryDate) Then
174                     dWork = CVDate(gtDeferredDeliveryTimeSettings.sDeliveryDate)
                    End If
                End If
            End If
  
176         lTotalSteps = (grdRecipients.Rows - 1) * lUpperLimit
  
178         If lTotalSteps > nStepSize Then
180             bMultipleStepJob = True
            End If
  
182         For lPeriodicSendDate = 1 To lUpperLimit
184             lStepCounter = lStepCounter + 1
186             cRecipientTemp.sRecipient = frmSMSMain.grdRecipients.TextMatrix(lRecipient, 1)   'Number of the Recipient
188             cRecipientTemp.sRecipientName = frmSMSMain.grdRecipients.TextMatrix(lRecipient, 0)
190             cRecipientTemp.sTransactionReferenceNumber = ConvertToASPSMSTimeFormat(Now()) & "_" & lStepCounter

192             If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime Then
194                 cRecipientTemp.sDeferredDeliveryTime = CStr(dWork)
                End If
    
196             If frmSMSMain.optSMSType(0).Value = True Then ' Text SMS
198                 cRecipientTemp.sMessageData = CreatePersonalizedMessageData(frmSMSMain.txtSMS.Text, lRecipient, bVariableMessageData)
                Else
                    'Do nothing Message will be taken from other place
                End If
    
200             If bVariableDeferredDeliveryTimeData = True Then

202                 Select Case gtDeferredDeliveryTimeSettings.nWaitingPeriod

                        Case 0 'Seconds
204                         dWork = DateAdd("s", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)
      
206                     Case 1 'Minutes
208                         dWork = DateAdd("n", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)

210                     Case 2 'Hours
212                         dWork = DateAdd("h", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)

214                     Case 3 'Days
216                         dWork = DateAdd("d", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)

218                     Case 4 'Weeks
220                         dWork = DateAdd("ww", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)

222                     Case 5 'Months
224                         dWork = DateAdd("m", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)

226                     Case Else
                            'Unexecpted, do nothing, probably an application bug
                    End Select
      
                End If
    
228             cRecipientData.Add Item:=cRecipientTemp
230             Set cRecipientTemp = Nothing

232             If lStepCounter / nStepSize = CLng(lStepCounter / nStepSize) Or lStepCounter = lTotalSteps Then
234                 bSuccess = ProcessSendAction(cRecipientData, bVariableMessageData, bVariableDeferredDeliveryTimeData, lJobId)
236                 Set cRecipientData = Nothing

238                 If bSuccess = False Then 'Error message already shown within ProcessSendAction
240                     bJobSuccessfullyProcessed = False
                        Exit Sub
                    End If

242                 If bMultipleStepJob Then
244                     Me.caption = LoadLanguageSpecificString(gnLanguage, 183) & " " & Trim(Str$(lStepCounter)) & " " & LoadLanguageSpecificString(gnLanguage, 184) & " " & Trim(Str$(lTotalSteps)) & " " & LoadLanguageSpecificString(gnLanguage, 185)

246                     DoEvents
                    End If
                End If

            Next

        Next

248     If bMultipleStepJob Then
250         Me.caption = LoadLanguageSpecificString(gnLanguage, 183) & " " & Trim(Str$(lTotalSteps)) & " " & LoadLanguageSpecificString(gnLanguage, 184) & " " & Trim(Str$(lTotalSteps)) & " " & LoadLanguageSpecificString(gnLanguage, 185)

252         DoEvents
        End If

254     If bJobSuccessfullyProcessed Then
256         MsgBox LoadLanguageSpecificString(gnLanguage, 180), vbInformation, gsApplicationName
        End If

258     ShowCredits False
        '<EhFooter>
        Exit Sub

cmdSend_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdSend_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdShowCredits_Click()
        '<EhHeader>
        On Error GoTo cmdShowCredits_Click_Err
        '</EhHeader>
100     ShowCredits True
        '<EhFooter>
        Exit Sub

cmdShowCredits_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdShowCredits_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdShowSendJournal_Click()
'        '<EhHeader>
'        On Error GoTo cmdShowSendJournal_Click_Err
'        '</EhHeader>
'100     Unload frmSendLog
'102     frmSendLog.Show
'        '<EhFooter>
'        Exit Sub
'
'cmdShowSendJournal_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmSMSMain.cmdShowSendJournal_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub cmdUpdate_Click05012003()
        '<EhHeader>
        On Error GoTo cmdUpdate_Click05012003_Err
        '</EhHeader>
        Dim lIDSavedRow As Long
        Dim lCurrentRow As Long
        Dim lMaxRows As Long
        Dim bFlagAddnew As Boolean

        On Error GoTo ErrorTrap

100     dtaPhonebook.Recordset.FindFirst "ID = " & mlCurrentPhoneBookID

102     If dtaPhonebook.Recordset.RecordCount = 0 Then
104         dtaPhonebook.Recordset.AddNew
106         bFlagAddnew = True
        Else
108         dtaPhonebook.Recordset.Edit
        End If

110     dtaPhonebook.Recordset(LoadLanguageSpecificString(gnLanguage, 196)) = txtPhonebookName.Text
112     dtaPhonebook.Recordset(LoadLanguageSpecificString(gnLanguage, 197)) = txtPhonebookPhoneNumber.Text
114     dtaPhonebook.Recordset(gsPhonebookVariableField(gnLanguage, 1)) = txtPhonebookVariableField(1).Text
116     dtaPhonebook.Recordset(gsPhonebookVariableField(gnLanguage, 2)) = txtPhonebookVariableField(2).Text
118     dtaPhonebook.Recordset(gsPhonebookVariableField(gnLanguage, 3)) = txtPhonebookVariableField(3).Text
120     dtaPhonebook.Recordset.UpDate

122     If bFlagAddnew Then
124         dtaPhonebook.Recordset.MoveLast
        End If

126     lIDSavedRow = dtaPhonebook.Recordset("ID")
128     dtaPhonebook.Refresh

130     lMaxRows = grdPhonebook.Rows

132     lCurrentRow = 0

134     Do While lCurrentRow <= lMaxRows And Val(grdPhonebook.TextMatrix(lCurrentRow, 0)) <> lIDSavedRow
136         lCurrentRow = lCurrentRow + 1

138         DoEvents
        Loop

140     grdPhonebook.TopRow = lCurrentRow

        Exit Sub
ErrorTrap:
        Exit Sub

        '<EhFooter>
        Exit Sub

cmdUpdate_Click05012003_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.cmdUpdate_Click05012003 " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdUpdate_Click()
'        '<EhHeader>
'        On Error GoTo cmdUpdate_Click_Err
'        '</EhHeader>
'        Dim rsMain As Recordset
'
'        Dim lIDSavedRow As Long
'        Dim lCurrentRow As Long
'        Dim lMaxRows As Long
'        Dim bFlagAddnew As Boolean
'
'        On Error GoTo ErrorTrap
'
'100     Set rsMain = gdbMain.OpenRecordset("select * from Phonebook where ID = " & Str$(mlCurrentPhoneBookID))
'
'102     If rsMain.RecordCount = 0 Then
'104         rsMain.AddNew
'106         bFlagAddnew = True
'        Else
'108         rsMain.Edit
'        End If
'
'110     rsMain("sName") = txtPhonebookName.Text
'112     rsMain("sNumber") = txtPhonebookPhoneNumber.Text
'114     rsMain("Variable1") = txtPhonebookVariableField(1).Text
'116     rsMain("Variable2") = txtPhonebookVariableField(2).Text
'118     rsMain("Variable3") = txtPhonebookVariableField(3).Text
'120     rsMain.UpDate
'
'122     If bFlagAddnew Then
'124         rsMain.MoveLast
'        End If
'
'126     lIDSavedRow = rsMain("ID")
'128     mlCurrentPhoneBookID = lIDSavedRow
'
'130     rsMain.Close
'
'132     dtaPhonebook.Refresh
'134     SetCorrectTopRowWithinPhonebook mlCurrentPhoneBookID
'136     dtaPhonebook.Refresh
'
'        Exit Sub
'ErrorTrap:
'        Exit Sub
'
'        '<EhFooter>
'        Exit Sub
'
'cmdUpdate_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmSMSMain.cmdUpdate_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     FormLoadWithoutSubClassing
102     gOldwndProcfrmSMSMain = GetWindowLong(TabMain.hWnd, GWL_WNDPROC)
104     SetWindowLong TabMain.hWnd, GWL_WNDPROC, AddressOf MouseWheelSupportWndProcfrmSMSMain

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     grsSendlog.Close
102     gdbMain.Close

104     SetWindowLong TabMain.hWnd, GWL_WNDPROC, gOldwndProcfrmSMSMain
106     Set frmSMSMain = Nothing

'108     End
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub grdPhonebook_Click()
        '<EhHeader>
        On Error GoTo grdPhonebook_Click_Err
        '</EhHeader>

100     If grdPhonebook.Rows <= 1 Then
            Exit Sub
        End If

102     mlCurrentPhoneBookID = grdPhonebook.TextMatrix(grdPhonebook.RowSel, 0)
104     dtaPhonebook.Recordset.FindFirst "ID = " & mlCurrentPhoneBookID
106     txtPhonebookName.Text = dtaPhonebook.Recordset(LoadLanguageSpecificString(gnLanguage, 196)) & ""
108     txtPhonebookPhoneNumber.Text = dtaPhonebook.Recordset(LoadLanguageSpecificString(gnLanguage, 197)) & ""
110     txtPhonebookVariableField(1).Text = dtaPhonebook.Recordset(gsPhonebookVariableField(gnLanguage, 1)) & ""
112     txtPhonebookVariableField(2).Text = dtaPhonebook.Recordset(gsPhonebookVariableField(gnLanguage, 2)) & ""
114     txtPhonebookVariableField(3).Text = dtaPhonebook.Recordset(gsPhonebookVariableField(gnLanguage, 3)) & ""

        '<EhFooter>
        Exit Sub

grdPhonebook_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.grdPhonebook_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub grdPhonebook_DblClick()
        '<EhHeader>
        On Error GoTo grdPhonebook_DblClick_Err
        '</EhHeader>
100     cmdAddRecipientsFromPhonebook_Click
        '<EhFooter>
        Exit Sub

grdPhonebook_DblClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.grdPhonebook_DblClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub scrlHorizontal_Change()
        '<EhHeader>
        On Error GoTo scrlHorizontal_Change_Err
        '</EhHeader>
100     scrlHorizontal_Scroll
        '<EhFooter>
        Exit Sub

scrlHorizontal_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.scrlHorizontal_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub scrlHorizontal_Scroll()
        '<EhHeader>
        On Error GoTo scrlHorizontal_Scroll_Err
        '</EhHeader>
100     imgPreviewWAPPushSMSPicture.Left = 0 - scrlHorizontal.Value
        '<EhFooter>
        Exit Sub

scrlHorizontal_Scroll_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.scrlHorizontal_Scroll " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub scrlVertical_Change()
        '<EhHeader>
        On Error GoTo scrlVertical_Change_Err
        '</EhHeader>
100     scrlVertical_Scroll
        '<EhFooter>
        Exit Sub

scrlVertical_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.scrlVertical_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub scrlVertical_Scroll()
        '<EhHeader>
        On Error GoTo scrlVertical_Scroll_Err
        '</EhHeader>
100     imgPreviewWAPPushSMSPicture.Top = 0 - scrlVertical.Value
        '<EhFooter>
        Exit Sub

scrlVertical_Scroll_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.scrlVertical_Scroll " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub TabMain_Click(PreviousTab As Integer)
        '<EhHeader>
        On Error GoTo TabMain_Click_Err
        '</EhHeader>
100     DebugPrint PreviousTab
102     If PreviousTab = 2 Then
104         txtSMSPreview.Text = txtSMS.Text
        End If
106     Form_Resize
        '<EhFooter>
        Exit Sub

TabMain_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.TabMain_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub TabMain_OLEDragDrop(Data As TabDlg.DataObject, _
                                Effect As Long, _
                                Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
        '<EhHeader>
        On Error GoTo TabMain_OLEDragDrop_Err
        '</EhHeader>
        Dim sFilePathAndNameOrig As String
        Dim sFilePathAndNameDestination As String
        Dim sFileNameOrig As String
        Dim sDescription As String

100     sFilePathAndNameOrig = FilesDropped(Data)

102     If Len(sFilePathAndNameOrig) = 0 Then
            Exit Sub
        End If

        'Truncate last 2 chars
104     sFilePathAndNameOrig = Left(sFilePathAndNameOrig, Len(sFilePathAndNameOrig) - 2)

106     sFileNameOrig = ExtractFilenameWithExtension(sFilePathAndNameOrig)
108     sDescription = ExtractFilenameWithoutExtension(sFilePathAndNameOrig)

110     CheckForRequiredDirectories App.Path & "\Images"

112     sFilePathAndNameDestination = App.Path & "\Images\" & CreateWAPPushDestinationFileName(sFileNameOrig)

114     FileCopy sFilePathAndNameOrig, sFilePathAndNameDestination

116     cmdDlgOpen.Filename = sFilePathAndNameDestination

118     Select Case True

            Case frmSMSMain.optSMSType(1).Value Or frmSMSMain.optSMSType(2).Value
120             txtPathLogo.Text = sFilePathAndNameDestination
    
122         Case frmSMSMain.optSMSType(3).Value
124             txtPathRingtone.Text = cmdDlgOpen.Filename
  
126         Case frmSMSMain.optSMSType(4).Value
128             txtPathPictureMessage.Text = sFilePathAndNameDestination
    
130         Case frmSMSMain.optSMSType(7).Value
132             txtWAPPushSMSPicturePath.Text = cmdDlgOpen.Filename
134             txtWAPPushSMSDescription.Text = sDescription
  
136         Case Else
                'Do nothing
  
        End Select

138     Form_Paint
        '<EhFooter>
        Exit Sub

TabMain_OLEDragDrop_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.TabMain_OLEDragDrop " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub TabMain_OLEDragOver(Data As TabDlg.DataObject, _
                                Effect As Long, _
                                Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single, _
                                State As Integer)
        '<EhHeader>
        On Error GoTo TabMain_OLEDragOver_Err
        '</EhHeader>
100     Me.Show
        '<EhFooter>
        Exit Sub

TabMain_OLEDragOver_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.TabMain_OLEDragOver " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtFilter_Change()
        '<EhHeader>
        On Error GoTo txtFilter_Change_Err
        '</EhHeader>
        Dim sTemp As String
        Dim sSQL As String

        On Error GoTo ErrorTrap

100     txtFilter.Enabled = False
102     Screen.MousePointer = vbHourglass

104     sSQL = PhonebookSQLStatement()

106     frmSMSMain.dtaPhonebook.RecordSource = sSQL
108     frmSMSMain.dtaPhonebook.Refresh

110     Screen.MousePointer = vbDefault
112     txtFilter.Enabled = True

        On Error Resume Next
114     txtFilter.SetFocus

        Exit Sub
ErrorTrap:

        Exit Sub
        '<EhFooter>
        Exit Sub

txtFilter_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.txtFilter_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtPictureMessageText_Change()
        '<EhHeader>
        On Error GoTo txtPictureMessageText_Change_Err
        '</EhHeader>
        Dim sTemp As String

100     sTemp = txtPictureMessageText.Text

        'BlinkingSMS not used
102     If 121 - Len(sTemp) >= 0 Then
104         lblCounterPictureMessage.caption = LoadLanguageSpecificString(gnLanguage, 181) & ": " & Trim(121 - Len(sTemp))
        Else
106         lblCounterPictureMessage.caption = LoadLanguageSpecificString(gnLanguage, 182)
        End If

        '<EhFooter>
        Exit Sub

txtPictureMessageText_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.txtPictureMessageText_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtRecipientDeliveryNotification_Change()
        '<EhHeader>
        On Error GoTo txtRecipientDeliveryNotification_Change_Err
        '</EhHeader>
100     RecipientDeliveryNotificationEventChange gnLanguage
        '<EhFooter>
        Exit Sub

txtRecipientDeliveryNotification_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.txtRecipientDeliveryNotification_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub txtSMS_Change()
        '<EhHeader>
        On Error GoTo txtSMS_Change_Err
        '</EhHeader>
        Dim sTemp As String
        Dim sOutput As String
        Dim sCaption As String
        Dim i As Integer
        Dim bBlinkingSMS As Boolean
        Dim bPlaceholdersUsed As Boolean

100     If chkBlinkingSMS.Value = 1 Then
102         bBlinkingSMS = True
        Else
104         bBlinkingSMS = False
        End If

106     sTemp = txtSMS.Text

108     If PersonalizedSMSModeUsed(sTemp) Then
110         lblCharsLeftCurrentMessage.caption = LoadLanguageSpecificString(gnLanguage, 221) & " " & LoadLanguageSpecificString(gnLanguage, 239)
112         lblNeededPartsCurrentMessage.caption = LoadLanguageSpecificString(gnLanguage, 222) & " " & LoadLanguageSpecificString(gnLanguage, 239)
            Exit Sub
        End If

        'VBCRLF is replaced by <Space>, because it only uses 1 character
114     Do While InStr(1, UCase(sTemp), vbCrLf) >= 1
116         sTemp = Replace(sTemp, vbCrLf, " ")
        Loop

118     If bBlinkingSMS Then    'Characters left is depending on using BlinkingSMS Feature or not

            'BlinkingSMS used
            'Set BlinkTag if not yet done
120         If InStr(1, UCase(sTemp), "<BLINK>") = 0 And InStr(1, UCase(sTemp), "</BLINK>") = 0 Then
122             sTemp = "<BLINK>" & sTemp
            End If

            '<BLINK> is replaced by <Space>, because it only uses 1 character
124         Do While InStr(1, UCase(sTemp), "<BLINK>") >= 1
126             sTemp = Replace(sTemp, "<BLINK>", " ", , , vbTextCompare)
            Loop

128         Do While InStr(1, UCase(sTemp), "</BLINK>") >= 1
130             sTemp = Replace(sTemp, "</BLINK>", " ", , , vbTextCompare)
            Loop

        End If

132     If NumberOfConcatenationParts(Len(sTemp), bBlinkingSMS) <= 9 Then
134         lblCharsLeftCurrentMessage.caption = LoadLanguageSpecificString(gnLanguage, 221) & " " & Trim(Str$(NumberOfRemainingCharsInLastConcatenationPart(Len(sTemp), bBlinkingSMS, True)))
136         lblNeededPartsCurrentMessage.caption = LoadLanguageSpecificString(gnLanguage, 222) & " " & Trim(Str$(NumberOfConcatenationParts(Len(sTemp), bBlinkingSMS)))
        Else
138         lblCharsLeftCurrentMessage.caption = LoadLanguageSpecificString(gnLanguage, 221) & " 0"
140         lblNeededPartsCurrentMessage.caption = LoadLanguageSpecificString(gnLanguage, 222) & " " & LoadLanguageSpecificString(gnLanguage, 482)
        End If

    

        '<EhFooter>
        Exit Sub

txtSMS_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.txtSMS_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtTemplateMessage_Change()
        '<EhHeader>
        On Error GoTo txtTemplateMessage_Change_Err
        '</EhHeader>
        Dim sTemp As String
100     sTemp = txtTemplateMessage.Text

102     If PersonalizedSMSModeUsed(sTemp) Then
104         lblCharsLeftTemplate.caption = LoadLanguageSpecificString(gnLanguage, 221) & " " & LoadLanguageSpecificString(gnLanguage, 239)
106         lblNeededPartsTemplate.caption = LoadLanguageSpecificString(gnLanguage, 222) & " " & LoadLanguageSpecificString(gnLanguage, 239)
            Exit Sub
        End If

108     lblCharsLeftTemplate.caption = LoadLanguageSpecificString(gnLanguage, 221) & " " & Trim(Str$(NumberOfRemainingCharsInLastConcatenationPart(Len(sTemp), False, False)))
110     lblNeededPartsTemplate.caption = LoadLanguageSpecificString(gnLanguage, 222) & " " & Trim(Str$(NumberOfConcatenationParts(Len(sTemp), False)))
        '<EhFooter>
        Exit Sub

txtTemplateMessage_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.txtTemplateMessage_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub txtUnicode_Change()
        '<EhHeader>
        On Error GoTo txtUnicode_Change_Err
        '</EhHeader>
        Dim sOrig As String
        Dim sTest1 As String
        Dim sTest2 As String
        Dim sOutput As String
        Dim sCaption As String
        Dim lTemp As Long
        Dim i As Integer
        Dim sUnicodeTest As String
        Dim sUnicodeConverted As String

        Dim nRemainingCharacters As Integer

100     sOrig = txtUnicode.Text

        'Copying
102     For i = 1 To Len(sOrig)
104         lTemp = AscW(Mid(sOrig, i, 1))
106         sUnicodeTest = sUnicodeTest & ChrW(lTemp)
        Next

108     sUnicodeTest = ""

        'VBCRLF is replaced by <Space>, because it only uses 1 character

        'In Codes umrechnen mit +/- Algorithmus
110     For i = 1 To Len(sOrig)
112         lTemp = AscW(Mid(sOrig, i, 1))

114         If lTemp < 0 Then
                'lTemp = (lTemp * -1) + 32767
116             lTemp = (Not lTemp) + 1
            End If

118         sUnicodeConverted = sUnicodeConverted & ChrW(lTemp)
120         sTest1 = sTest1 & lTemp & ", "
        Next

122     For i = 1 To Len(sUnicodeConverted)
124         lTemp = AscW(Mid(sUnicodeConverted, i, 1))

126         If lTemp < 0 Then
128             lTemp = (lTemp * -1) + 32767
            End If

130         sTest2 = sTest2 & lTemp & ", "
        Next

132     nRemainingCharacters = 70 - Len(sOrig)

134     If nRemainingCharacters >= 0 Then
136         lblCharsLeftCurrentMessageUnicode.caption = LoadLanguageSpecificString(gnLanguage, 486) & " " & Trim(Str$(nRemainingCharacters))
        Else
138         lblCharsLeftCurrentMessageUnicode.caption = LoadLanguageSpecificString(gnLanguage, 486) & " " & LoadLanguageSpecificString(gnLanguage, 482)
        End If

        '<EhFooter>
        Exit Sub

txtUnicode_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.txtUnicode_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub winsckLogoControl_Connect()
        '<EhHeader>
        On Error GoTo winsckLogoControl_Connect_Err
        '</EhHeader>
        On Error GoTo ErrorTrap
        Dim sRessource As String
        Dim sTemp As String
        Dim nRandomLogo As Integer
        Dim sRandomLogoFilename As String
        Dim sUserAgent As String
        Dim sHost As String

        'IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE
        'It's allowed to:
        '- Use the code of this sub within a 3rd party application based on ASPSMS
        '- Use the logos of ASPSMS / Handylogos unlimited within a 3rd party application based on ASPSMS

        'It's not allowed to:
        '- Use the code of this sub within a 3rd party application NOT based on ASPSMS
        '- Use the logos of ASPSMS / Handylogos unlimited within a 3rd party application NOT based on ASPSMS
        '- Download the logos of ASPSMS / Handylogos unlimited and provide them on other websites / printmedias /
        '  in any form to any parties
        'ANY COPYRIGHT VIOLATIONS WILL BE PROSECUTED BY SERIOUS TERMS OF CIVIL AND CRIMINAL LAW
        'IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE
        '
        'If used in a ASPSMS 3rd Party Application, please adapt the value of Useragent to your
        'applicationname and version THANK YOU!
100     VersionSpecificAction 38, , , sUserAgent

102     VersionSpecificAction 39, , , sRessource

104     VersionSpecificAction 40, , , sHost
        '
106     sTemp = "GET " & sRessource & " HTTP/1.1" & Chr(13) & Chr(10)
108     sTemp = sTemp & "Accept: */*" & Chr(13) & Chr(10)
110     sTemp = sTemp & "User-Agent: " & sUserAgent & Chr(13) & Chr(10)
112     sTemp = sTemp & "Host: " & sHost & Chr(13) & Chr(10)
114     sTemp = sTemp & Chr(13) & Chr(10)

116     winsckLogoControl.SendData sTemp

        Exit Sub
ErrorTrap:
        Exit Sub

        '<EhFooter>
        Exit Sub

winsckLogoControl_Connect_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.winsckLogoControl_Connect " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub winsckLogoControl_DataArrival(ByVal bytesTotal As Long)
        '<EhHeader>
        On Error GoTo winsckLogoControl_DataArrival_Err
        '</EhHeader>
        Dim sTemp As String
        Dim sHost As String
        Dim bSuccess As Boolean
        Dim i As Integer

        On Error GoTo ErrorTrap:

100     VersionSpecificAction 40, , , sHost

102     winsckLogoControl.GetData sTemp
104     msControlFileNumberOfRandomLogos = msControlFileNumberOfRandomLogos & sTemp

106     If Len(msControlFileNumberOfRandomLogos) > 200 Then
108         mlNumberOfRandomLogos = ParseReceivedLogoControlData(msControlFileNumberOfRandomLogos, bSuccess)
  
110         If bSuccess Then

112             For i = 0 To 20
114                 msReceivedLogoData(i) = ""

116                 With winsckRandomLogo(i)
118                     .Close
120                     .LocalPort = 0
122                     .connect sHost, 80

124                     DoEvents
                    End With

                Next

            End If

        End If

        Exit Sub
ErrorTrap:
        Exit Sub
        '<EhFooter>
        Exit Sub

winsckLogoControl_DataArrival_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.winsckLogoControl_DataArrival " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub winsckRandomLogo_Connect(Index As Integer)
        '<EhHeader>
        On Error GoTo winsckRandomLogo_Connect_Err
        '</EhHeader>
        On Error GoTo ErrorTrap
        Dim sResource As String
        Dim sTemp As String
        Dim nRandomLogo As Integer
        Dim sRandomLogoFilename As String
        Dim sUserAgent As String
        Dim sHost As String
        Dim sLogoPath As String

        'IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE
        'It's allowed to:
        '- Use the code of this sub within a 3rd party application based on ASPSMS
        '- Use the logos of ASPSMS / Handylogos unlimited within a 3rd party application based on ASPSMS

        'It's not allowed to:
        '- Use the code of this sub within a 3rd party application NOT based on ASPSMS
        '- Use the logos of ASPSMS / Handylogos unlimited within a 3rd party application NOT based on ASPSMS
        '- Download the logos of ASPSMS / Handylogos unlimited and provide them on other websites / printmedias /
        '  in any form to any parties
        'ANY COPYRIGHT VIOLATIONS WILL BE PROSECUTED BY SERIOUS TERMS OF CIVIL AND CRIMINAL LAW
        'IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE / IMPORTANT COPYRIGHT NOTICE

        'If used in a ASPSMS 3rd Party Application, please adapt the value of Useragent to your
        'applicationname and version THANK YOU!
        '
        'lstLogoDebug.AddItem "winsckRandomLogo CONNECT Index: " & Index

100     VersionSpecificAction 40, , , sHost
102     VersionSpecificAction 38, , , sUserAgent
104     VersionSpecificAction 41, , , sLogoPath

106     Randomize CLng(Second(Now) + Minute(Now))
108     nRandomLogo = Int(Rnd * mlNumberOfRandomLogos + 1)
110     sRandomLogoFilename = Right("000" & Trim(Str(nRandomLogo)), 4) & ".bmp"
112     sResource = sLogoPath & sRandomLogoFilename

114     sTemp = "GET " & sResource & " HTTP/1.1" & Chr(13) & Chr(10)
116     sTemp = sTemp & "Accept: */*" & Chr(13) & Chr(10)
118     sTemp = sTemp & "User-Agent: " & sUserAgent & Chr(13) & Chr(10)
120     sTemp = sTemp & "Host: " & sHost & Chr(13) & Chr(10)
122     sTemp = sTemp & Chr(13) & Chr(10)

124     winsckRandomLogo(Index).SendData sTemp

        Exit Sub
ErrorTrap:
        Exit Sub

        '<EhFooter>
        Exit Sub

winsckRandomLogo_Connect_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.winsckRandomLogo_Connect " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub winsckRandomLogo_DataArrival(Index As Integer, _
                                         ByVal bytesTotal As Long)
        '<EhHeader>
        On Error GoTo winsckRandomLogo_DataArrival_Err
        '</EhHeader>
        Dim sTemp As String
        Dim sPicture As String
        Dim sFilename As String
        Dim bSuccess As Boolean

        On Error GoTo ErrorTrap:

100     winsckRandomLogo(Index).GetData sTemp
102     msReceivedLogoData(Index) = msReceivedLogoData(Index) & sTemp

104     If Len(msReceivedLogoData(Index)) > 300 Then

106         sPicture = ParseReceivedPictureData(msReceivedLogoData(Index), bSuccess)

108         If bSuccess Then
110             ProcessRandomLogoDisplayUpdate Index, sPicture

112             sFilename = App.Path & "\" & Right("0000" & Trim(Str$(Index)), 4) & ".bmp"
114             imgRandomLogo(Index).Picture = LoadPicture(sFilename)

116             If Index = 0 Then
118                 imgRandomLogo_Click 0
                End If
            End If
        End If

        Exit Sub
ErrorTrap:
        Exit Sub
        '<EhFooter>
        Exit Sub

winsckRandomLogo_DataArrival_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.winsckRandomLogo_DataArrival " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub CheckCredit()
        '<EhHeader>
        On Error GoTo CheckCredit_Err
        '</EhHeader>
        Dim SMS As Booster
        Dim rCredits As Double
        Dim sMessage As String

100     Set SMS = New Booster

102     SMS.Hosts = App.Path & "\smshosts.txt"
104     SMS.UserKey = frmSMSMain.txtUserkey.Text
106     SMS.Password = frmSMSMain.txtPassword.Text

108     Screen.MousePointer = vbHourglass

110     rCredits = SMS.Credits

112     sMessage = "Your Current balance is: " & Trim(Str$(rCredits)) & " units. Your identifier is set as: " & txtOriginator.Text

114     Screen.MousePointer = vbDefault

116     Set SMS = Nothing

118     lblCurrentBalance.caption = sMessage

        '<EhFooter>
        Exit Sub

CheckCredit_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSMSMain.CheckCredit " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
