VERSION 5.00
Begin VB.Form frmFilterSendlog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter"
   ClientHeight    =   2520
   ClientLeft      =   4515
   ClientTop       =   4275
   ClientWidth     =   5895
   Icon            =   "frmFilterSendlog.frx":0000
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
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100 Unload Me
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmFilterSendlog.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOk_Click()
        '<EhHeader>
        On Error GoTo cmdOk_Click_Err
        '</EhHeader>

100 If cboRecipientName.ListIndex <> -1 Then
102   If cboRecipientName.ItemData(cboRecipientName.ListIndex) <> gcNoFilter Then
104     gtSendLogFilterSettings.bRecipientNameFilterUsed = True
106     gtSendLogFilterSettings.sRecipientName = cboRecipientName.List(cboRecipientName.ListIndex)
      Else
108     gtSendLogFilterSettings.bRecipientNameFilterUsed = False
110     gtSendLogFilterSettings.sRecipientName = ""
      End If
    End If

112 If cboRecipient.ListIndex <> -1 Then
114   If cboRecipient.ItemData(cboRecipient.ListIndex) <> gcNoFilter Then
116     gtSendLogFilterSettings.bRecipientFilterUsed = True
118     gtSendLogFilterSettings.sRecipient = cboRecipient.List(cboRecipient.ListIndex)
      Else
120     gtSendLogFilterSettings.bRecipientFilterUsed = False
122     gtSendLogFilterSettings.sRecipient = cboRecipient.List(cboRecipient.ListIndex)
      End If
    End If

124 If cboDeliveryStatus.ListIndex <> -1 Then
126   Select Case cboDeliveryStatus.ItemData(cboDeliveryStatus.ListIndex)
  
      Case gcNoFilter
128   gtSendLogFilterSettings.bDeliveryStatusFilterUsed = False
130   gtSendLogFilterSettings.nDeliveryStatus = cboDeliveryStatus.ItemData(cboDeliveryStatus.ListIndex)
    
132   Case gcQuestionMarks
134   gtSendLogFilterSettings.bDeliveryStatusFilterUsed = True
136   gtSendLogFilterSettings.nDeliveryStatus = cboDeliveryStatus.ItemData(cboDeliveryStatus.ListIndex)

138   Case Else
140   gtSendLogFilterSettings.bDeliveryStatusFilterUsed = True
142   gtSendLogFilterSettings.nDeliveryStatus = cboDeliveryStatus.ItemData(cboDeliveryStatus.ListIndex)
  
      End Select

    End If

144 If cboReasonCode.ListIndex <> -1 Then
146   Select Case cboReasonCode.ItemData(cboReasonCode.ListIndex)
      Case gcNoFilter
148   gtSendLogFilterSettings.bReasonCodeFilterUsed = False
150   gtSendLogFilterSettings.nReasonCode = cboReasonCode.ItemData(cboReasonCode.ListIndex)
  
152   Case gcQuestionMarks
154   gtSendLogFilterSettings.bReasonCodeFilterUsed = False
156   gtSendLogFilterSettings.nReasonCode = cboReasonCode.ItemData(cboReasonCode.ListIndex)
  
158   Case Else
160   gtSendLogFilterSettings.bReasonCodeFilterUsed = True
162   gtSendLogFilterSettings.nReasonCode = cboReasonCode.ItemData(cboReasonCode.ListIndex)
  
      End Select
  
    End If

164 Call frmSendLog.FormLoadWithoutSubClassing
166 Unload Me
        '<EhFooter>
        Exit Sub

cmdOk_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmFilterSendlog.cmdOk_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    Dim sSQL As String
100 CenterForm Me
102 Screen.MousePointer = vbHourglass
    On Error GoTo ErrorTrap:
104 AdjustFontControls Me
106 AdjustLanguageSettings gnLanguage
108 dtaFilter.DatabaseName = gsPathAndDatabaseName
110 dtaFilter.RecordSource = gsSQLSendjournal
112 dtaFilter.Refresh

114 If gtSendLogFilterSettings.bRecipientNameFilterUsed Then
116   SetListIndexWithText cboRecipientName, gtSendLogFilterSettings.sRecipientName
    Else
118     SetListIndexWithItemData cboRecipientName, gcNoFilter
    End If

120 If gtSendLogFilterSettings.bRecipientFilterUsed Then
122   SetListIndexWithText cboRecipient, gtSendLogFilterSettings.sRecipient
    Else
124   SetListIndexWithItemData cboRecipient, gcNoFilter
    End If

126 If gtSendLogFilterSettings.bDeliveryStatusFilterUsed Then
128   SetListIndexWithItemData cboDeliveryStatus, gtSendLogFilterSettings.nDeliveryStatus
    Else
130   SetListIndexWithItemData cboDeliveryStatus, gcNoFilter
    End If

132 If gtSendLogFilterSettings.bReasonCodeFilterUsed Then
134   SetListIndexWithItemData cboReasonCode, gtSendLogFilterSettings.nReasonCode
    Else
136   SetListIndexWithItemData cboReasonCode, gcNoFilter
    End If
138 Screen.MousePointer = vbDefault

    Exit Sub
ErrorTrap:
140 Screen.MousePointer = vbDefault
    Exit Sub

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmFilterSendlog.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Public Sub AdjustLanguageSettings(nLanguage As Integer)
        '<EhHeader>
        On Error GoTo AdjustLanguageSettings_Err
        '</EhHeader>
100 dtaFilter.DatabaseName = gsPathAndDatabaseName

102 lblName.Caption = LoadLanguageSpecificString(nLanguage, 401)
104 lblRecipient.Caption = LoadLanguageSpecificString(nLanguage, 402)
106 lblDeliveryStatus.Caption = LoadLanguageSpecificString(nLanguage, 403)
108 lblReasonCode.Caption = LoadLanguageSpecificString(nLanguage, 404)

110 cmdOK.Caption = LoadLanguageSpecificString(nLanguage, 1)
112 cmdCancel.Caption = LoadLanguageSpecificString(nLanguage, 2)
      
    'Recipient Names
114 dtaFilter.RecordSource = "select sName from SendJournal group by sName order by sName"
116 dtaFilter.Refresh
118 dtaFilter.Recordset.MoveFirst
120 cboRecipientName.Clear
122 cboRecipientName.AddItem LoadLanguageSpecificString(nLanguage, 321)
124 cboRecipientName.ItemData(cboRecipientName.NewIndex) = gcNoFilter
126 Do While Not dtaFilter.Recordset.EOF
128   cboRecipientName.AddItem dtaFilter.Recordset("sName") & ""
130   dtaFilter.Recordset.MoveNext
    Loop

    'Recipient Phonenumbers
132 dtaFilter.RecordSource = "select sRecipient from SendJournal group by sRecipient order by sRecipient"
134 dtaFilter.Refresh
136 dtaFilter.Recordset.MoveFirst
138 cboRecipient.Clear
140 cboRecipient.AddItem LoadLanguageSpecificString(nLanguage, 321)
142 cboRecipient.ItemData(cboRecipient.NewIndex) = gcNoFilter
144 Do While Not dtaFilter.Recordset.EOF
146   cboRecipient.AddItem dtaFilter.Recordset("sRecipient") & ""
148   dtaFilter.Recordset.MoveNext
    Loop

    'Deliverystatus
150 cboDeliveryStatus.Clear
152 cboDeliveryStatus.AddItem LoadLanguageSpecificString(nLanguage, 321)
154 cboDeliveryStatus.ItemData(cboDeliveryStatus.NewIndex) = gcNoFilter
156 cboDeliveryStatus.AddItem "???"
158 cboDeliveryStatus.ItemData(cboDeliveryStatus.NewIndex) = gcQuestionMarks
160 cboDeliveryStatus.AddItem "-1 " & DeliveryStatusFromDeliveryStatusCode("-1")
162 cboDeliveryStatus.ItemData(cboDeliveryStatus.NewIndex) = -1
164 cboDeliveryStatus.AddItem " 0 " & DeliveryStatusFromDeliveryStatusCode("0")
166 cboDeliveryStatus.ItemData(cboDeliveryStatus.NewIndex) = 0
168 cboDeliveryStatus.AddItem " 1 " & DeliveryStatusFromDeliveryStatusCode("1")
170 cboDeliveryStatus.ItemData(cboDeliveryStatus.NewIndex) = 1
172 cboDeliveryStatus.AddItem " 2 " & DeliveryStatusFromDeliveryStatusCode("2")
174 cboDeliveryStatus.ItemData(cboDeliveryStatus.NewIndex) = 2

    'Reasoncode
176 cboReasonCode.Clear
178 cboReasonCode.AddItem LoadLanguageSpecificString(nLanguage, 321)
180 cboReasonCode.ItemData(cboReasonCode.NewIndex) = gcNoFilter
182 cboReasonCode.AddItem "???"
184 cboReasonCode.ItemData(cboReasonCode.NewIndex) = gcQuestionMarks
186 cboReasonCode.AddItem "000 " & ReasonFromReasonCode("000"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 0
188 cboReasonCode.AddItem "001 " & ReasonFromReasonCode("001"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 1
190 cboReasonCode.AddItem "002 " & ReasonFromReasonCode("002"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 2
192 cboReasonCode.AddItem "003 " & ReasonFromReasonCode("003"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 3
194 cboReasonCode.AddItem "004 " & ReasonFromReasonCode("004"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 4
196 cboReasonCode.AddItem "005 " & ReasonFromReasonCode("005"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 5
198 cboReasonCode.AddItem "006 " & ReasonFromReasonCode("006"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 6
200 cboReasonCode.AddItem "007 " & ReasonFromReasonCode("007"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 7
202 cboReasonCode.AddItem "008 " & ReasonFromReasonCode("008"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 8
204 cboReasonCode.AddItem "009 " & ReasonFromReasonCode("009"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 9
206 cboReasonCode.AddItem "010 " & ReasonFromReasonCode("010"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 10
208 cboReasonCode.AddItem "100 " & ReasonFromReasonCode("100"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 100
210 cboReasonCode.AddItem "101 " & ReasonFromReasonCode("101"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 101
212 cboReasonCode.AddItem "102 " & ReasonFromReasonCode("102"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 102
214 cboReasonCode.AddItem "103 " & ReasonFromReasonCode("103"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 103
216 cboReasonCode.AddItem "104 " & ReasonFromReasonCode("104"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 104
218 cboReasonCode.AddItem "105 " & ReasonFromReasonCode("105"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 105
220 cboReasonCode.AddItem "106 " & ReasonFromReasonCode("106"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 106
222 cboReasonCode.AddItem "107 " & ReasonFromReasonCode("107"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 107
224 cboReasonCode.AddItem "108 " & ReasonFromReasonCode("108"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 108
226 cboReasonCode.AddItem "109 " & ReasonFromReasonCode("109"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 109
228 cboReasonCode.AddItem "110 " & ReasonFromReasonCode("110"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 110
230 cboReasonCode.AddItem "111 " & ReasonFromReasonCode("111"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 111
232 cboReasonCode.AddItem "112 " & ReasonFromReasonCode("112"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 112
234 cboReasonCode.AddItem "113 " & ReasonFromReasonCode("113"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 113
236 cboReasonCode.AddItem "114 " & ReasonFromReasonCode("114"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 114
238 cboReasonCode.AddItem "115 " & ReasonFromReasonCode("115"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 115
240 cboReasonCode.AddItem "116 " & ReasonFromReasonCode("116"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 116
242 cboReasonCode.AddItem "117 " & ReasonFromReasonCode("117"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 117
244 cboReasonCode.AddItem "118 " & ReasonFromReasonCode("118"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 118
246 cboReasonCode.AddItem "119 " & ReasonFromReasonCode("119"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 119
248 cboReasonCode.AddItem "120 " & ReasonFromReasonCode("120"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 120
250 cboReasonCode.AddItem "121 " & ReasonFromReasonCode("121"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 121
252 cboReasonCode.AddItem "122 " & ReasonFromReasonCode("122"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 122
254 cboReasonCode.AddItem "123 " & ReasonFromReasonCode("123"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 123
256 cboReasonCode.AddItem "124 " & ReasonFromReasonCode("124"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 124
258 cboReasonCode.AddItem "125 " & ReasonFromReasonCode("125"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 125
260 cboReasonCode.AddItem "126 " & ReasonFromReasonCode("126"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 126
262 cboReasonCode.AddItem "127 " & ReasonFromReasonCode("127"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 127
264 cboReasonCode.AddItem "200 " & ReasonFromReasonCode("200"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 200
266 cboReasonCode.AddItem "201 " & ReasonFromReasonCode("201"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 201
268 cboReasonCode.AddItem "202 " & ReasonFromReasonCode("202"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 202
270 cboReasonCode.AddItem "203 " & ReasonFromReasonCode("203"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 203
272 cboReasonCode.AddItem "204 " & ReasonFromReasonCode("204"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 204
274 cboReasonCode.AddItem "205 " & ReasonFromReasonCode("205"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 205
276 cboReasonCode.AddItem "206 " & ReasonFromReasonCode("206"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 206
278 cboReasonCode.AddItem "207 " & ReasonFromReasonCode("207"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 207
280 cboReasonCode.AddItem "208 " & ReasonFromReasonCode("208"): cboReasonCode.ItemData(cboReasonCode.NewIndex) = 208
        '<EhFooter>
        Exit Sub

AdjustLanguageSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmFilterSendlog.AdjustLanguageSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
