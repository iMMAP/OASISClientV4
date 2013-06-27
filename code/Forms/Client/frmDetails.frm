VERSION 5.00
Begin VB.Form frmDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Details"
   ClientHeight    =   8760
   ClientLeft      =   4545
   ClientTop       =   1500
   ClientWidth     =   5895
   Icon            =   "frmDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRemarks 
      Caption         =   "fraRemarks"
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   5655
      Begin VB.Label lblDetailsRemarks 
         Caption         =   "lblDetailsRemarks"
         Height          =   2775
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame fraTransmissionDetails 
      Caption         =   "fraTransmissionDetails"
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   5655
      Begin VB.Label lblDeferredDeliveryTime 
         Caption         =   "lblDeferredDeliveryTime"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label lblSubmissionDate 
         Caption         =   "lblSubmissionDate"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblNotificationDate 
         Caption         =   "lblNotificationDate"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblDeliveryStatus 
         Caption         =   "lblDeliveryStatus"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblDeliveryStatusText 
         Caption         =   "lblDeliveryStatusText"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblReasonCode 
         Caption         =   "lblReasonCode"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblReasonText 
         Caption         =   "lblReasonText"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblNotificationDateBuffered 
         Caption         =   "lblNotificationDateBuffered"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblReasonCodeBuffered 
         Caption         =   "lblReasonCodeBuffered"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblReasonTextBuffered 
         Caption         =   "lblReasonTextBuffered"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblDetailsDeferredDeliveryTime 
         Caption         =   "lblDetailsDeferredDeliveryTime"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label lblDetailsSubmissionDate 
         Caption         =   "lblDetailsSubmissionDate"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblDetailsNotificationDate 
         Caption         =   "lblDetailsNotificationDate"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblDetailsDeliveryStatus 
         Caption         =   "lblDetailsDeliveryStatus"
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblDetailsDeliveryStatusText 
         Caption         =   "lblDetailsDeliveryStatusText"
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label lblDetailsReasonCode 
         Caption         =   "lblDetailsReasonCode"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label lblDetailsReasonText 
         Caption         =   "lblDetailsReasonText"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label lblDetailsNotificationDateBuffered 
         Caption         =   "lblDetailsNotificationDateBuffered"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label lblDetailsReasonCodeBuffered 
         Caption         =   "lblDetailsReasonCodeBuffered"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblDetailsReasonTextBuffered 
         Caption         =   "lblDetailsReasonTextBuffered"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label lblTransferTime 
         Caption         =   "lbTransfertime"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblDetailsTransferTime 
         Caption         =   "lblDetailsTransferTime"
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   2640
         Width           =   3135
      End
   End
   Begin VB.Frame fraRecipientAndMessage 
      Caption         =   "fraRecipientAndMessage"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      Begin VB.Label lblName 
         Caption         =   "lblName"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblRecipient 
         Caption         =   "lblRecipient"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblMessage 
         Caption         =   "lblMessage"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblDetailsName 
         Caption         =   "lblDetailsName"
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblDetailsRecipient 
         Caption         =   "lblDetailsRecipient"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblDetailsMessage 
         Caption         =   "lblDetailsMessage"
         Height          =   855
         Left            =   2280
         TabIndex        =   4
         Top             =   720
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   8280
      Width           =   1815
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub PrepareDisplay(lRecordID As Long, frmParent As Form)
        '<EhHeader>
        On Error GoTo PrepareDisplay_Err
        '</EhHeader>
    Dim rsMain As Recordset
    Dim sRemarks As String
    Dim sSQL As String

100 sSQL = "select IIf([dNotificationDate]-[dSubmissionDate]>3,'> 72h',IIf([dNotificationDate]-[dSubmissionDate]>2,'> 48h',IIf([dNotificationDate]-[dSubmissionDate]>1,'> 24h',IIf([dSubmissionDate]<>0 And [dNotificationDate]<>0,CVDate([dNotificationDate]-[dSubmissionDate]),Null)))) AS TransferTime, * from SendJournal where lID = " & Str$(lRecordID)

102 Set rsMain = gdbMain.OpenRecordset(sSQL)

104 If rsMain.BOF And rsMain.EOF Then
      'Do nothing
    Else
106   lblDetailsName.Caption = rsMain("sName") & ""
108   lblDetailsRecipient.Caption = rsMain("sRecipient") & ""
110   lblDetailsMessage.Caption = rsMain("sMessage") & ""
112   lblDetailsDeferredDeliveryTime.Caption = rsMain("dDeferredDeliveryTime") & ""
114   lblDetailsSubmissionDate.Caption = rsMain("dSubmissionDate") & ""
116   lblDetailsNotificationDate.Caption = rsMain("dNotificationDate") & ""
118   lblDetailsDeliveryStatus.Caption = rsMain("sDeliveryStatus") & ""
120   lblDetailsDeliveryStatusText.Caption = rsMain("sDeliveryStatusText") & ""
122   lblDetailsReasonCode.Caption = rsMain("sReasonCode") & ""
124   lblDetailsReasonText.Caption = rsMain("sReasonText") & ""
126   lblDetailsNotificationDateBuffered.Caption = rsMain("dNotificationDateBuffered") & ""
128   lblDetailsReasonCodeBuffered.Caption = rsMain("sReasonCodeBuffered") & ""
130   lblDetailsReasonTextBuffered.Caption = rsMain("sReasonTextBuffered") & ""
132   lblDetailsTransferTime.Caption = rsMain("TransferTime") & ""
    End If

134 Select Case True
      Case rsMain("sDeliveryStatus") & "" = "0" And rsMain("sReasonCodeBuffered") & "" = "" 'Immediately Delivered
136   sRemarks = LoadLanguageSpecificString(gnLanguage, 366)
  
138   Case rsMain("sDeliveryStatus") & "" = "0" And rsMain("sReasonCodeBuffered") & "" <> "" 'Delayed Delivered
140   sRemarks = LoadLanguageSpecificString(gnLanguage, 368) & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 373) & vbCrLf & _
                 ReasonRemarkFromReasonCode(rsMain("sReasonCodeBuffered") & "", rsMain("sDeliveryStatus") & "") & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 369)
  
142   Case rsMain("sDeliveryStatus") & "" = "1"
144   sRemarks = LoadLanguageSpecificString(gnLanguage, 367) & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 373) & vbCrLf & _
                 ReasonRemarkFromReasonCode(rsMain("sReasonCode") & "", rsMain("sDeliveryStatus") & "")
  
146   Case rsMain("sDeliveryStatus") & "" = "2" And rsMain("sReasonCodeBuffered") & "" = "" 'Immediately Failed
148   sRemarks = LoadLanguageSpecificString(gnLanguage, 372) & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 373) & vbCrLf & _
                 ReasonRemarkFromReasonCode(rsMain("sReasonCode"), rsMain("sDeliveryStatus"))
  
150   Case rsMain("sDeliveryStatus") & "" = "2" And rsMain("sReasonCodeBuffered") & "" <> "" And _
                 (rsMain("sReasonCodeBuffered") & "" <> rsMain("sReasonCode") & "") 'Delayed and Failed because of other reason
152   sRemarks = LoadLanguageSpecificString(gnLanguage, 368) & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 373) & " " & _
                 rsMain("sReasonTextBuffered") & ", Code: " & rsMain("sReasonCodeBuffered") & "" & vbCrLf & _
                 ReasonRemarkFromReasonCode(rsMain("sReasonCodeBuffered"), rsMain("sDeliveryStatus")) & vbCrLf & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 371) & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 373) & " " & _
                 rsMain("sReasonText") & ", Code " & rsMain("sReasonCode") & "" & vbCrLf & _
                 ReasonRemarkFromReasonCode(rsMain("sReasonCode"), rsMain("sDeliveryStatus"))
  
154   Case rsMain("sDeliveryStatus") & "" = "2" And rsMain("sReasonCodeBuffered") & "" <> "" And _
                 (rsMain("sReasonCodeBuffered") & "" = rsMain("sReasonCode") & "") 'Delayed and Failed because of timeout
156   sRemarks = LoadLanguageSpecificString(gnLanguage, 368) & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 373) & " " & _
                 rsMain("sReasonTextBuffered") & ", Code: " & rsMain("sReasonCodeBuffered") & "" & vbCrLf & _
                 ReasonRemarkFromReasonCode(rsMain("sReasonCodeBuffered"), rsMain("sDeliveryStatus")) & vbCrLf & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 371) & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 373) & " " & _
                 rsMain("sReasonText") & ", Code " & rsMain("sReasonCode") & "" & vbCrLf & _
                 ReasonRemarkFromReasonCode(rsMain("sReasonCode"), rsMain("sDeliveryStatus"))
158   sRemarks = LoadLanguageSpecificString(gnLanguage, 368) & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 373) & " " & _
                 rsMain("sReasonTextBuffered") & ", Code: " & rsMain("sReasonCodeBuffered") & "" & vbCrLf & _
                 ReasonRemarkFromReasonCode(rsMain("sReasonCodeBuffered"), rsMain("sDeliveryStatus")) & vbCrLf & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 371) & vbCrLf & _
                 LoadLanguageSpecificString(gnLanguage, 373) & " " & _
                 vbCrLf & _
                 ReasonRemarkFromReasonCode("108", "2")
    End Select

160 lblDetailsRemarks.Caption = sRemarks

162 rsMain.Close

164 frmDetails.Show vbModal, frmParent
        '<EhFooter>
        Exit Sub

PrepareDisplay_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDetails.PrepareDisplay " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Function ReasonTextTillDELETE(sReasonCode As String) As String
        '<EhHeader>
        On Error GoTo ReasonTextTillDELETE_Err
        '</EhHeader>
100 Select Case sReasonCode

      Case ""
102   ReasonTextTillDELETE = ""

104   Case "000"
106   ReasonTextTillDELETE = "Unbekannter Empfänger"

108   Case "001"
110   ReasonTextTillDELETE = "Vorübergehend kein Service"

112   Case "002"
114   ReasonTextTillDELETE = "Vorübergehend kein Service"

116   Case "003"
118   ReasonTextTillDELETE = "Vorübergehend kein Service"

120   Case "004"
122   ReasonTextTillDELETE = "Vorübergehend kein Service"

124   Case "005"
126   ReasonTextTillDELETE = "Vorübergehend kein Service"

128   Case "006"
130   ReasonTextTillDELETE = "Vorübergehend kein Service"

132   Case "007"
134   ReasonTextTillDELETE = "Vorübergehend kein Service"

136   Case "008"
138   ReasonTextTillDELETE = "Vorübergehend kein Service"

140   Case "009"
142   ReasonTextTillDELETE = "Unbekannter Fehler"

144   Case "010"
146   ReasonTextTillDELETE = "Netzwerk Timeout"

148   Case "100"
150   ReasonTextTillDELETE = "Dienst nicht unterstützt"

152   Case "101"
154   ReasonTextTillDELETE = "Unbekannter Empfänger"

156   Case "102"
158   ReasonTextTillDELETE = "Dienst nicht verfügbar"

160   Case "103"
162   ReasonTextTillDELETE = "Anrufsperre"

164   Case "104"
166   ReasonTextTillDELETE = "Operation gesperrt"

168   Case "105"
170   ReasonTextTillDELETE = "Service Center überlastet"

172   Case "106"
174   ReasonTextTillDELETE = "Dienst nicht unterstützt"

176   Case "107"
178   ReasonTextTillDELETE = "Empfänger vorübergehend nicht erreichbar"

180   Case "108"
182   ReasonTextTillDELETE = "Auslieferungsfehler"

184   Case "109"
186   ReasonTextTillDELETE = "Service Center überlastet"

188   Case "110"
190   ReasonTextTillDELETE = "Protokollfehler"

192   Case "111"
194   ReasonTextTillDELETE = "Mobiltelefon des Empfängers ohne SMS"

196   Case "112"
198   ReasonTextTillDELETE = "Unbekanntes Service Center"

200   Case "113"
202   ReasonTextTillDELETE = "Service Center überlastet"

204   Case "114"
206   ReasonTextTillDELETE = "Illegales Mobiltelefon des Empfängers"

208   Case "115"
210   ReasonTextTillDELETE = "Empfänger kein Kunde"

212   Case "116"
214   ReasonTextTillDELETE = "Fehler im Mobiltelefon des Empfängers"

216   Case "117"
218   ReasonTextTillDELETE = "Untere Protokollschicht für SMS nicht verfügbar"

220   Case "118"
222   ReasonTextTillDELETE = "Systemfehler"

224   Case "119"
226   ReasonTextTillDELETE = "PLMN Systemfehler"

228   Case "120"
230   ReasonTextTillDELETE = "HLR Systemfehler"

232   Case "121"
234   ReasonTextTillDELETE = "VLR Systemfehler"

236   Case "122"
238   ReasonTextTillDELETE = "Vorangegangener VLR Systemfehler"

240   Case "123"
242   ReasonTextTillDELETE = "Steuer-MSC Systemfehler"

244   Case "124"
246   ReasonTextTillDELETE = "VMSC Systemfehler"

248   Case "125"
250   ReasonTextTillDELETE = "EIR Systemfehler"

252   Case "126"
254   ReasonTextTillDELETE = "Systemfehler"

256   Case "127"
258   ReasonTextTillDELETE = "Unerwartete Daten"

260   Case "200"
262   ReasonTextTillDELETE = "Fehler bei der Adressierung des Service Centers"

264   Case "201"
266   ReasonTextTillDELETE = "Ungültige absolute Speicherzeit"

268   Case "202"
270   ReasonTextTillDELETE = "Nachricht grösser als Maximum"

272   Case "203"
274   ReasonTextTillDELETE = "GSM-Nachricht kann nicht ausgepackt werden"

276   Case "204"
278   ReasonTextTillDELETE = "Übersetzung in IA5 ALPHABET nicht möglich"

280   Case "205"
282   ReasonTextTillDELETE = "Ungültiges Format der Speicherzeit"

284   Case "206"
286   ReasonTextTillDELETE = "Ungültige Empfängeradresse"

288   Case "207"
290   ReasonTextTillDELETE = "Nachricht zweimal gesendet"

292   Case "208"
294   ReasonTextTillDELETE = "Ungültiger Nachrichtentyp"

296   Case Else
298   ReasonTextTillDELETE = "Unbekannter Fehlercode"

    End Select

        '<EhFooter>
        Exit Function

ReasonTextTillDELETE_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDetails.ReasonTextTillDELETE " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdOk_Click()
        '<EhHeader>
        On Error GoTo cmdOk_Click_Err
        '</EhHeader>
100 Unload Me
        '<EhFooter>
        Exit Sub

cmdOk_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDetails.cmdOk_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100 CenterForm Me
102 AdjustFontControls Me
104 AdjustLanguageSettings gnLanguage
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDetails.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Public Sub AdjustLanguageSettings(nLanguage As Integer)
        '<EhHeader>
        On Error GoTo AdjustLanguageSettings_Err
        '</EhHeader>
100 fraRecipientAndMessage.Caption = LoadLanguageSpecificString(nLanguage, 381)
102 fraTransmissionDetails.Caption = LoadLanguageSpecificString(nLanguage, 382)
104 fraRemarks.Caption = LoadLanguageSpecificString(nLanguage, 383)
106 lblName.Caption = LoadLanguageSpecificString(nLanguage, 384)
108 lblRecipient.Caption = LoadLanguageSpecificString(nLanguage, 385)
110 lblMessage.Caption = LoadLanguageSpecificString(nLanguage, 386)
112 lblSubmissionDate.Caption = LoadLanguageSpecificString(nLanguage, 387)
114 lblNotificationDate.Caption = LoadLanguageSpecificString(nLanguage, 388)
116 lblDeliveryStatus.Caption = LoadLanguageSpecificString(nLanguage, 389)
118 lblDeliveryStatusText.Caption = LoadLanguageSpecificString(nLanguage, 390)
120 lblReasonCode.Caption = LoadLanguageSpecificString(nLanguage, 391)
122 lblReasonText.Caption = LoadLanguageSpecificString(nLanguage, 392)
124 lblNotificationDateBuffered.Caption = LoadLanguageSpecificString(nLanguage, 393)
126 lblReasonCodeBuffered.Caption = LoadLanguageSpecificString(nLanguage, 394)
128 lblReasonTextBuffered.Caption = LoadLanguageSpecificString(nLanguage, 395)
130 lblDeferredDeliveryTime.Caption = LoadLanguageSpecificString(nLanguage, 396)
132 lblTransferTime.Caption = LoadLanguageSpecificString(nLanguage, 397)
        '<EhFooter>
        Exit Sub

AdjustLanguageSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDetails.AdjustLanguageSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub lblDetailsMessage_Click()
        '<EhHeader>
        On Error GoTo lblDetailsMessage_Click_Err
        '</EhHeader>
100 MsgBox lblDetailsMessage.Caption, vbInformation, gsApplicationName
        '<EhFooter>
        Exit Sub

lblDetailsMessage_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDetails.lblDetailsMessage_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


