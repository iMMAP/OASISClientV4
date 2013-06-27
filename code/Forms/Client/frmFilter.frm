VERSION 5.00
Begin VB.Form frmFilterSendlog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter"
   ClientHeight    =   2520
   ClientLeft      =   4515
   ClientTop       =   4275
   ClientWidth     =   5895
   Icon            =   "frmFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox cboReasonCode 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.ComboBox cboDeliveryStatus 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.ComboBox cboRecipient 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.ComboBox cboRecipientName 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Data dtaFilter 
      Caption         =   "dtaFilter"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblReasonCode 
      Caption         =   "lblReasonCode"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblDeliveryStatus 
      Caption         =   "lblDeliveryStatus"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblRecipient 
      Caption         =   "lblRecipient"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblName 
      Caption         =   "lblName"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmFilterSendlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()

'If cboRecipientName.ListIndex <> -1 Then
'  If cboRecipientName.ItemData(cboRecipientName.ListIndex) <> gcNoFilter Then
'    gtSendLogFilterSettings.bRecipientNameFilterUsed = True
'    gtSendLogFilterSettings.sRecipientName = cboRecipientName.List(cboRecipientName.ListIndex)
'  Else
'    gtSendLogFilterSettings.bRecipientNameFilterUsed = False
'    gtSendLogFilterSettings.sRecipientName = ""
'  End If
'End If
'
'If cboRecipient.ListIndex <> -1 Then
'  If cboRecipient.ItemData(cboRecipient.ListIndex) <> gcNoFilter Then
'    gtSendLogFilterSettings.bRecipientFilterUsed = True
'    gtSendLogFilterSettings.sRecipient = cboRecipient.List(cboRecipient.ListIndex)
'  Else
'    gtSendLogFilterSettings.bRecipientFilterUsed = False
'    gtSendLogFilterSettings.sRecipient = cboRecipient.List(cboRecipient.ListIndex)
'  End If
'End If
'
'If cboDeliveryStatus.ListIndex <> -1 Then
'  If cboDeliveryStatus.ItemData(cboDeliveryStatus.ListIndex) <> gcNoFilter Then
'    gtSendLogFilterSettings.bDeliveryStatusFilterUsed = True
'    gtSendLogFilterSettings.nDeliveryStatus = cboDeliveryStatus.ItemData(cboDeliveryStatus.ListIndex)
'  Else
'    gtSendLogFilterSettings.bDeliveryStatusFilterUsed = False
'    gtSendLogFilterSettings.nDeliveryStatus = cboDeliveryStatus.ItemData(cboDeliveryStatus.ListIndex)
'  End If
'End If
'
'If cboReasonCode.ListIndex <> -1 Then
'  If cboReasonCode.ItemData(cboReasonCode.ListIndex) <> gcNoFilter Then
'    gtSendLogFilterSettings.bReasonCodeFilterUsed = True
'    gtSendLogFilterSettings.nReasonCode = cboReasonCode.ItemData(cboReasonCode.ListIndex)
'  Else
'    gtSendLogFilterSettings.bReasonCodeFilterUsed = False
'    gtSendLogFilterSettings.nReasonCode = cboReasonCode.ItemData(cboReasonCode.ListIndex)
'  End If
'End If
'
'Call frmSendLog.Form_Load
'Unload Me
End Sub


Private Sub Form_Load()
'Dim sSQL As String
'CenterForm Me
'On Error GoTo ErrorTrap:
'AdjustFontControls Me
'AdjustLanguageSettings gnLanguage
'dtaFilter.DatabaseName = App.Path & gsDatabaseName
'dtaFilter.RecordSource = gsSQLSendjournal
'dtaFilter.Refresh
'
'If gtSendLogFilterSettings.bRecipientNameFilterUsed Then
'  cboRecipientName.Text = gtSendLogFilterSettings.sRecipientName
'Else
'    SetListIndexWithItemData cboRecipientName, gcNoFilter
'End If
'
'If gtSendLogFilterSettings.bRecipientFilterUsed Then
'  cboRecipient.Text = gtSendLogFilterSettings.sRecipient
'Else
'  SetListIndexWithItemData cboRecipient, gcNoFilter
'End If
'
'If gtSendLogFilterSettings.bDeliveryStatusFilterUsed Then
'  SetListIndexWithItemData cboDeliveryStatus, gtSendLogFilterSettings.nDeliveryStatus
'Else
'  SetListIndexWithItemData cboDeliveryStatus, gcNoFilter
'End If
'
'If gtSendLogFilterSettings.bReasonCodeFilterUsed Then
'  SetListIndexWithItemData cboReasonCode, gtSendLogFilterSettings.nReasonCode
'Else
'  SetListIndexWithItemData cboReasonCode, gcNoFilter
'End If
'
'Exit Sub
'ErrorTrap:
'Exit Sub

End Sub
Public Sub AdjustLanguageSettings(nLanguage As Integer)
''Me.Caption = LoadLanguageSpecificString(nLanguage, 41)
''cmdInquireDeliveryNotification.Caption = LoadLanguageSpecificString(nLanguage, 42)
''cmdDeleteRows.Caption = LoadLanguageSpecificString(nLanguage, 43)
''cmdClose.Caption = LoadLanguageSpecificString(nLanguage, 44)
'
''cboWaitingPeriod.AddItem LoadLanguageSpecificString(nLanguage, 187)
'
'dtaFilter.DatabaseName = App.Path & gsDatabaseName
'
'lblName.caption = LoadLanguageSpecificString(nLanguage, 401)
'lblRecipient.caption = LoadLanguageSpecificString(nLanguage, 402)
'lblDeliveryStatus.caption = LoadLanguageSpecificString(nLanguage, 403)
'lblReasonCode.caption = LoadLanguageSpecificString(nLanguage, 404)
'
'cmdOk.caption = LoadLanguageSpecificString(nLanguage, 1)
'cmdCancel.caption = LoadLanguageSpecificString(nLanguage, 2)
'
''Recipient Names
'dtaFilter.RecordSource = "select sName from SendJournal group by sName order by sName"
'dtaFilter.Refresh
'dtaFilter.Recordset.MoveFirst
'cboRecipientName.Clear
'cboRecipientName.AddItem LoadLanguageSpecificString(nLanguage, 321)
'cboRecipientName.ItemData(cboRecipientName.NewIndex) = gcNoFilter
'Do While Not dtaFilter.Recordset.EOF
'  cboRecipientName.AddItem dtaFilter.Recordset("sName") & ""
'  'MsgBox dtaFilter.Recordset("Name") & ""
'  dtaFilter.Recordset.MoveNext
'Loop
'
''Recipient Phonenumbers
'dtaFilter.RecordSource = "select sRecipient from SendJournal group by sRecipient order by sRecipient"
'dtaFilter.Refresh
'dtaFilter.Recordset.MoveFirst
'cboRecipient.Clear
'cboRecipient.AddItem LoadLanguageSpecificString(nLanguage, 321)
'cboRecipient.ItemData(cboRecipient.NewIndex) = gcNoFilter
'Do While Not dtaFilter.Recordset.EOF
'  cboRecipient.AddItem dtaFilter.Recordset("sRecipient") & ""
'  'MsgBox dtaFilter.Recordset("Recipient") & ""
'  dtaFilter.Recordset.MoveNext
'Loop
'
''Deliverystatus
'cboDeliveryStatus.Clear
'cboDeliveryStatus.AddItem LoadLanguageSpecificString(nLanguage, 321)
'cboDeliveryStatus.ItemData(cboDeliveryStatus.NewIndex) = gcNoFilter
'cboDeliveryStatus.AddItem "-1 " & DeliveryStatusFromDeliveryStatusCode("-1")
'cboDeliveryStatus.ItemData(cboDeliveryStatus.NewIndex) = -1
'cboDeliveryStatus.AddItem " 0 " & DeliveryStatusFromDeliveryStatusCode("0")
'cboDeliveryStatus.ItemData(cboDeliveryStatus.NewIndex) = 0
'cboDeliveryStatus.AddItem " 1 " & DeliveryStatusFromDeliveryStatusCode("1")
'cboDeliveryStatus.ItemData(cboDeliveryStatus.NewIndex) = 1
'cboDeliveryStatus.AddItem " 2 " & DeliveryStatusFromDeliveryStatusCode("2")
'cboDeliveryStatus.ItemData(cboDeliveryStatus.NewIndex) = 2
'
''Reasoncode
'cboReasonCode.Clear
'cboReasonCode.AddItem LoadLanguageSpecificString(nLanguage, 321)
'cboReasonCode.ItemData(cboReasonCode.NewIndex) = gcNoFilter
'cboReasonCode.AddItem "000 " & ReasonFromReasonCode("000"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 0
'cboReasonCode.AddItem "001 " & ReasonFromReasonCode("001"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 1
'cboReasonCode.AddItem "002 " & ReasonFromReasonCode("002"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 2
'cboReasonCode.AddItem "003 " & ReasonFromReasonCode("003"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 3
'cboReasonCode.AddItem "004 " & ReasonFromReasonCode("004"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 4
'cboReasonCode.AddItem "005 " & ReasonFromReasonCode("005"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 5
'cboReasonCode.AddItem "006 " & ReasonFromReasonCode("006"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 6
'cboReasonCode.AddItem "007 " & ReasonFromReasonCode("007"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 7
'cboReasonCode.AddItem "008 " & ReasonFromReasonCode("008"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 8
'cboReasonCode.AddItem "009 " & ReasonFromReasonCode("009"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 9
'cboReasonCode.AddItem "010 " & ReasonFromReasonCode("010"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 10
'cboReasonCode.AddItem "100 " & ReasonFromReasonCode("100"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 100
'cboReasonCode.AddItem "101 " & ReasonFromReasonCode("101"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 101
'cboReasonCode.AddItem "102 " & ReasonFromReasonCode("102"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 102
'cboReasonCode.AddItem "103 " & ReasonFromReasonCode("103"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 103
'cboReasonCode.AddItem "104 " & ReasonFromReasonCode("104"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 104
'cboReasonCode.AddItem "105 " & ReasonFromReasonCode("105"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 105
'cboReasonCode.AddItem "106 " & ReasonFromReasonCode("106"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 106
'cboReasonCode.AddItem "107 " & ReasonFromReasonCode("107"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 107
'cboReasonCode.AddItem "108 " & ReasonFromReasonCode("108"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 108
'cboReasonCode.AddItem "109 " & ReasonFromReasonCode("109"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 109
'cboReasonCode.AddItem "110 " & ReasonFromReasonCode("110"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 110
'cboReasonCode.AddItem "111 " & ReasonFromReasonCode("111"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 111
'cboReasonCode.AddItem "112 " & ReasonFromReasonCode("112"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 112
'cboReasonCode.AddItem "113 " & ReasonFromReasonCode("113"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 113
'cboReasonCode.AddItem "114 " & ReasonFromReasonCode("114"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 114
'cboReasonCode.AddItem "115 " & ReasonFromReasonCode("115"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 115
'cboReasonCode.AddItem "116 " & ReasonFromReasonCode("116"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 116
'cboReasonCode.AddItem "117 " & ReasonFromReasonCode("117"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 117
'cboReasonCode.AddItem "118 " & ReasonFromReasonCode("118"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 118
'cboReasonCode.AddItem "119 " & ReasonFromReasonCode("119"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 119
'cboReasonCode.AddItem "120 " & ReasonFromReasonCode("120"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 120
'cboReasonCode.AddItem "121 " & ReasonFromReasonCode("121"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 121
'cboReasonCode.AddItem "122 " & ReasonFromReasonCode("122"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 122
'cboReasonCode.AddItem "123 " & ReasonFromReasonCode("123"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 123
'cboReasonCode.AddItem "124 " & ReasonFromReasonCode("124"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 124
'cboReasonCode.AddItem "125 " & ReasonFromReasonCode("125"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 125
'cboReasonCode.AddItem "126 " & ReasonFromReasonCode("126"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 126
'cboReasonCode.AddItem "127 " & ReasonFromReasonCode("127"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 127
'cboReasonCode.AddItem "200 " & ReasonFromReasonCode("200"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 200
'cboReasonCode.AddItem "201 " & ReasonFromReasonCode("201"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 201
'cboReasonCode.AddItem "202 " & ReasonFromReasonCode("202"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 202
'cboReasonCode.AddItem "203 " & ReasonFromReasonCode("203"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 203
'cboReasonCode.AddItem "204 " & ReasonFromReasonCode("204"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 204
'cboReasonCode.AddItem "205 " & ReasonFromReasonCode("205"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 205
'cboReasonCode.AddItem "206 " & ReasonFromReasonCode("206"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 206
'cboReasonCode.AddItem "207 " & ReasonFromReasonCode("207"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 207
'cboReasonCode.AddItem "208 " & ReasonFromReasonCode("208"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 208
End Sub
