VERSION 5.00
Begin VB.Form frmDeliveryNotifications 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delivery Notifications"
   ClientHeight    =   6600
   ClientLeft      =   3645
   ClientTop       =   1920
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton cmdInquireDeliveryNotifications 
      Caption         =   "Inquire Delivery Notifications"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Data datSendJournal 
      Caption         =   "datSendJournal"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000E&
      Height          =   5835
      Left            =   120
      ScaleHeight     =   5775
      ScaleWidth      =   7830
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7890
      Begin VB.VScrollBar VScrollPersonGroupDetails 
         Height          =   5775
         LargeChange     =   10
         Left            =   7560
         Max             =   1
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Value           =   1
         Width           =   255
      End
      Begin VB.CheckBox chkSMS 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "chkSMS"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   0
         Value           =   1  'Checked
         Width           =   7455
      End
   End
End
Attribute VB_Name = "frmDeliveryNotifications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub LoadPendingNotifications()
Dim i As Long
Dim sCaption As String
Dim sMessage As String
Const cPersonGroupDetails = 25
Const cSelectionGroupDetails = 25
Dim sSQL As String

'sSQL = "SELECT ID, Name, Recipient, TransactionReferenceNumber, Message, DeferredDeliveryTime, " & _
'       "IIf([SubmissionDate] Is Not Null,[SubmissionDate],'???') AS Submission, IIf([NotificationDate] Is Not Null, " & _
'       "[NotificationDate],'???') AS Notification, IIf([NotificationDate] Is Null And [SubmissionDate] Is Null,'???', " & _
'       "IIf([NotificationDate]-[SubmissionDate]>3,'> 72h',IIf([NotificationDate]-[SubmissionDate]>2,'> 48h', " & _
'       "IIf([NotificationDate]-[SubmissionDate]>1,'> 24h',IIf([SubmissionDate]<>0 And [NotificationDate]<>0, " & _
'       "CVDate([NotificationDate]-[SubmissionDate]),Null))))) AS TransferTime, DeliveryStatus, DeliveryStatusText, " & _
'       "ReasonCode, ReasonText, NotificationDateBuffered, ReasonCodeBuffered, ReasonTextBuffered " & _
'       "From Sendjournal"
       
       
sSQL = "SELECT ID, Name, Recipient, TransactionReferenceNumber, Message, DeferredDeliveryTime " & _
       "From Sendjournal where Deliverystatus is null or (Deliverystatus is not null and DeferredDeliveryTime <= cvdate('" & Now() & "')) " & _
       "order by ID"
'MsgBox sSQL
       
       
datSendJournal.RecordSource = sSQL
'For calculating correct numbers of records
datSendJournal.Refresh
If datSendJournal.Recordset.BOF = False And datSendJournal.Recordset.EOF = False Then
  datSendJournal.Recordset.MoveLast
  datSendJournal.Recordset.MoveFirst
End If
MsgBox datSendJournal.Recordset.RecordCount
Me.datSendJournal.Recordset.MoveFirst
i = 0
Do While Not frmDeliveryNotifications.datSendJournal.Recordset.EOF
  If i <> 0 Then  'Because not yet loaded
    Load frmDeliveryNotifications.chkSMS(i)
  End If
  
  sCaption = ""
  If datSendJournal.Recordset("Name") <> "" Then
    sCaption = datSendJournal.Recordset("Name") & ", "
  End If
  
  If datSendJournal.Recordset("Recipient") <> "" Then
    sCaption = sCaption & datSendJournal.Recordset("Recipient") & ", "
  End If
  
  If datSendJournal.Recordset("Message") <> "" Then
    sMessage = datSendJournal.Recordset("Message") & ""
    If Len(sMessage) >= 40 Then
      sMessage = Left(sMessage, 40) & "..."
    End If
    sCaption = sCaption & sMessage & ", "
  End If
  
  If datSendJournal.Recordset("DeferredDeliveryTime") <> "" Then
    sCaption = sCaption & datSendJournal.Recordset("DeferredDeliveryTime") & ", "
  End If
  
  'Truncate last ", "
  If Len(sCaption) >= 2 Then
    sCaption = Left(sCaption, Len(sCaption) - 2)
  End If
    
  sCaption = Replace(sCaption, vbCrLf, " ")
  'sCaption = Replace(sCaption, vbCr, " ")
  'sCaption = Replace(sCaption, vbLf, " ")
  
  Me.chkSMS(i).caption = sCaption
  Me.chkSMS(i).toolTipText = (datSendJournal.Recordset("Message") & "")
  'List1.AddItem (sCaption & vbCrLf & datSendJournal.Recordset("Message") & "")
  Me.chkSMS(i).Tag = Me.datSendJournal.Recordset("ID")
  Me.chkSMS(i).Left = 60
  Me.chkSMS(i).Top = i * 240
  Me.chkSMS(i).Visible = True
  
  
  'frmSMSBlaster.cboPhonebookFilter.AddItem frmSMSBlaster.datGroups.Recordset("Name") & ""
  'frmSMSBlaster.cboPhonebookFilter.ItemData(frmSMSBlaster.cboPhonebookFilter.NewIndex) = frmSMSBlaster.datGroups.Recordset("ID")
  
  i = i + 1
  Me.datSendJournal.Recordset.MoveNext
Loop

Me.VScrollPersonGroupDetails.Max = Me.datSendJournal.Recordset.RecordCount - cPersonGroupDetails
Me.VScrollPersonGroupDetails.Min = 0
Me.VScrollPersonGroupDetails.Value = 0
'Stop
'frmFilterRecipients.VScrollSelectionGroupDetails.Max = frmSMSBlaster.datGroups.Recordset.RecordCount - cSelectionGroupDetails + 1
'frmFilterRecipients.VScrollSelectionGroupDetails.Min = 0
'frmFilterRecipients.VScrollSelectionGroupDetails.Value = 0
'frmFilterRecipients.chkSelectionGroup(0).Value = 1

'frmSMSBlaster.cboPhonebookFilter.ListIndex = 0
End Sub

Private Sub cmdHelp_Click()
Dim sHelp As String
'sHelp = "Empfangsbestätigungen stellen Ihnen wichtige Informationen über den Auslieferungsstatus "
'sHelp = sHelp & "Ihrer Mitteilungen zur Verfügung. " & vbCrLf
sHelp = "Der Versand von SMS Mitteilungen kann aus vielerlei Gründen fehlschlagen oder sich verzögern: " & vbCrLf
sHelp = sHelp & "Nicht eingeschaltete oder gesperrte Handys, kein Empfang, unbekannte Nummern, technische Defekte im terminierenden Netzwerk, usw. " & vbCrLf
sHelp = sHelp & "Wir empfehlen aus diesen Gründen, nach jedem Versand die Empfangsbestätigungen für jede SMS abzufragen. " & vbCrLf & vbCrLf
sHelp = sHelp & "Falls Sie jeweils grössere SMS Mengen verschicken, bieten Ihnen Empfangsbestätigungen zusätzlich eine einfache und effektive Möglichkeit, "
sHelp = sHelp & "die Qualität Ihrer Handynummern zu prüfen. Handynummern, an die der Versand fehlgeschlagen ist, können damit identifiziert und von weiteren Versendungen ausgeschlossen werden." & vbCrLf & vbCrLf
sHelp = sHelp & "Kosten:" & vbCrLf
sHelp = sHelp & "Die Abfrage einer Empfangsbestätigung für eine bestimmte Mitteilung kann jeweils dreimal gratis vorgenommen werden, ab dem vierten Mal kostet jede weitere Abfrage 0.25 Credits. "
'sHelp = sHelp & ""
'sHelp = sHelp & "Für SMS, die nicht direkt nach dem Versand ausgeliefert wurden konnten, wird versucht, diese während rund 24 Stunden doch noch auszuliefern. "
'sHelp = sHelp & "Nach Ablauf dieser Zeitspanne verändern sich die Empfangsbestätigungen nicht mehr, der Versand ist dann geglückt oder definitiv fehlgeschlagen."
        
MsgBox sHelp, vbInformation, gsApplicationName


'Dim i As Integer
'sHelp = ""
'For i = 1 To 3000
'  sHelp = sHelp & Str$(i)
'Next
'MsgBox sHelp
End Sub


Private Sub Form_Load()
'Screen.MousePointer = vbHourglass
'AdjustFontControls Me
'AdjustLanguageSettings gnLanguage
'datSendJournal.DatabaseName = App.Path & gsDatabaseName
'LoadPendingNotifications
'Screen.MousePointer = vbDefault
End Sub


Sub ScrollGroups(vScrollInput As VScrollBar)
Dim i As Integer
Dim nBegin As Integer
Dim nCurrentGroup As Integer

Select Case True
  Case vScrollInput Is Me.VScrollPersonGroupDetails
  nBegin = Me.VScrollPersonGroupDetails.Value
  For i = Me.chkSMS.LBound To Me.chkSMS.UBound
    If i >= nBegin And (i - nBegin) < 25 Then
      Me.chkSMS(i).Visible = True
      Me.chkSMS(i).Enabled = True
      Me.chkSMS(i).Left = 60
      Me.chkSMS(i).Top = (i - nBegin) * 240
    Else
      Me.chkSMS(i).Visible = False
      Me.chkSMS(i).Enabled = False
    End If
  Next
  
'  Case vScrollInput Is Me.VScrollSelectionGroupDetails
'  nBegin = Me.VScrollSelectionGroupDetails.Value
'  If nBegin = 0 Then Stop
'  Stop
'  For i = Me.chkSelectionGroup.LBound To Me.chkSelectionGroup.ubound
'    If i >= nBegin And (i - nBegin) < 17 Then
'      Me.chkSelectionGroup(i).Visible = True
'      Me.chkSelectionGroup(i).Enabled = True
'      Me.chkSelectionGroup(i).Left = 60
'      Me.chkSelectionGroup(i).Top = (i - nBegin) * 240
'    Else
'      Me.chkSelectionGroup(i).Visible = False
'      Me.chkSelectionGroup(i).Enabled = False
'    End If
'  Next

  Case Else
  MsgBox "Case else"
  
End Select


End Sub

Public Sub AdjustLanguageSettings(nLanguage As Integer)
Me.caption = LoadLanguageSpecificString(nLanguage, 21)
'fraRecipientAndMessage.Caption = LoadLanguageSpecificString(nLanguage, 381)
'fraTransmissionDetails.Caption = LoadLanguageSpecificString(nLanguage, 382)
'fraRemarks.Caption = LoadLanguageSpecificString(nLanguage, 383)
'lblName.Caption = LoadLanguageSpecificString(nLanguage, 384)
'lblRecipient.Caption = LoadLanguageSpecificString(nLanguage, 385)
'lblMessage.Caption = LoadLanguageSpecificString(nLanguage, 386)
'lblSubmissionDate.Caption = LoadLanguageSpecificString(nLanguage, 387)
'lblNotificationDate.Caption = LoadLanguageSpecificString(nLanguage, 388)
'lblDeliveryStatus.Caption = LoadLanguageSpecificString(nLanguage, 389)
'lblDeliveryStatusText.Caption = LoadLanguageSpecificString(nLanguage, 390)
'lblReasonCode.Caption = LoadLanguageSpecificString(nLanguage, 391)
'lblReasonText.Caption = LoadLanguageSpecificString(nLanguage, 392)
'lblNotificationDateBuffered.Caption = LoadLanguageSpecificString(nLanguage, 393)
'lblReasonCodeBuffered.Caption = LoadLanguageSpecificString(nLanguage, 394)
'lblReasonTextBuffered.Caption = LoadLanguageSpecificString(nLanguage, 395)
'lblDeferredDeliveryTime.Caption = LoadLanguageSpecificString(nLanguage, 396)
'lblTransferTime.Caption = LoadLanguageSpecificString(nLanguage, 397)
End Sub

Private Sub VScrollPersonGroupDetails_Change()
ScrollGroups VScrollPersonGroupDetails
End Sub


Private Sub VScrollPersonGroupDetails_Scroll()
ScrollGroups VScrollPersonGroupDetails
End Sub


