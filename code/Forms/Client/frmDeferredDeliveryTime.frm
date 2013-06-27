VERSION 5.00
Begin VB.Form frmDeferredDeliveryTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings Deferred Delivery Time"
   ClientHeight    =   6615
   ClientLeft      =   1305
   ClientTop       =   1230
   ClientWidth     =   5880
   Icon            =   "frmDeferredDeliveryTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTimeZone 
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame fraSchedulingPreview 
      Caption         =   "Scheduling Preview"
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   5655
      Begin VB.ListBox lstSchedulingPreview 
         Height          =   2400
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.CheckBox chkUseDeferredDeliveryTime 
      Caption         =   "Use Deferred Delivery Time"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.OptionButton optPeriodicSMS 
      Caption         =   "Periodic SMS"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.OptionButton optSingleSMS 
      Caption         =   "Single SMS"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   -1  'True
      Width           =   2535
   End
   Begin VB.Frame fraPeriodicSMS 
      Caption         =   "Periodic SMS"
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   5655
      Begin VB.ComboBox cboWaitingPeriod 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtWaitingTime 
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Text            =   "txtWaitingTime"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtNumberOfMessages 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Text            =   "txtNumberOfMessages"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtStartingDate 
         Height          =   285
         Left            =   3120
         TabIndex        =   11
         Text            =   "txtStartingDate"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblWaitingTime 
         Caption         =   "Waiting time between messages"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblNumberOfMessages 
         Caption         =   "Number of messages"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblStartingDate 
         Caption         =   "Starting Date"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraSingleSMS 
      Caption         =   "Single SMS"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   5655
      Begin VB.TextBox txtDeferredDeliveryTime 
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Text            =   "txtDeferredDeliveryTime"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblDeliveryDate 
         Caption         =   "Delivery Date"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Label lblTimezone 
      Caption         =   "Timezone"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   2535
   End
End
Attribute VB_Name = "frmDeferredDeliveryTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AdjustLanguageSettings(nLanguage As Integer)
        '<EhHeader>
        On Error GoTo AdjustLanguageSettings_Err
        '</EhHeader>
100     Me.Caption = LoadLanguageSpecificString(nLanguage, 11)
102     cmdOK.Caption = LoadLanguageSpecificString(nLanguage, 1)
104     cmdCancel.Caption = LoadLanguageSpecificString(nLanguage, 2)
106     chkUseDeferredDeliveryTime.Caption = LoadLanguageSpecificString(nLanguage, 12)
108     optSingleSMS.Caption = LoadLanguageSpecificString(nLanguage, 13)
110     optPeriodicSMS.Caption = LoadLanguageSpecificString(nLanguage, 14)
112     fraSingleSMS.Caption = LoadLanguageSpecificString(nLanguage, 13)
114     lblDeliveryDate.Caption = LoadLanguageSpecificString(nLanguage, 15)
116     fraPeriodicSMS.Caption = LoadLanguageSpecificString(nLanguage, 14)
118     lblStartingDate.Caption = LoadLanguageSpecificString(nLanguage, 16)
120     lblNumberOfMessages.Caption = LoadLanguageSpecificString(nLanguage, 17)
122     lblWaitingTime.Caption = LoadLanguageSpecificString(nLanguage, 18)
124     fraSchedulingPreview.Caption = LoadLanguageSpecificString(nLanguage, 19)
126     lblTimezone.Caption = LoadLanguageSpecificString(nLanguage, 212)

128     cboWaitingPeriod.AddItem LoadLanguageSpecificString(nLanguage, 187)
130     cboWaitingPeriod.AddItem LoadLanguageSpecificString(nLanguage, 188)
132     cboWaitingPeriod.AddItem LoadLanguageSpecificString(nLanguage, 189)
134     cboWaitingPeriod.AddItem LoadLanguageSpecificString(nLanguage, 190)
136     cboWaitingPeriod.AddItem LoadLanguageSpecificString(nLanguage, 191)
138     cboWaitingPeriod.AddItem LoadLanguageSpecificString(nLanguage, 192)

140     cboTimeZone.AddItem "GMT -12:00"
142     cboTimeZone.AddItem "GMT -11:00"
144     cboTimeZone.AddItem "GMT -10:00"
146     cboTimeZone.AddItem "GMT -09:00"
148     cboTimeZone.AddItem "GMT -08:00"
150     cboTimeZone.AddItem "GMT -07:00"
152     cboTimeZone.AddItem "GMT -06:00"
154     cboTimeZone.AddItem "GMT -05:00"
156     cboTimeZone.AddItem "GMT -04:00"
158     cboTimeZone.AddItem "GMT -03:00"
160     cboTimeZone.AddItem "GMT -02:00"
162     cboTimeZone.AddItem "GMT -01:00"
164     cboTimeZone.AddItem "GMT -00:00"
166     cboTimeZone.AddItem "GMT +01:00"
168     cboTimeZone.AddItem "GMT +02:00"
170     cboTimeZone.AddItem "GMT +03:00"
172     cboTimeZone.AddItem "GMT +04:00"
174     cboTimeZone.AddItem "GMT +05:00"
176     cboTimeZone.AddItem "GMT +06:00"
178     cboTimeZone.AddItem "GMT +07:00"
180     cboTimeZone.AddItem "GMT +08:00"
182     cboTimeZone.AddItem "GMT +09:00"
184     cboTimeZone.AddItem "GMT +10:00"
186     cboTimeZone.AddItem "GMT +11:00"
188     cboTimeZone.AddItem "GMT +12:00"
190     cboTimeZone.AddItem "GMT +13:00"
        '<EhFooter>
        Exit Sub

AdjustLanguageSettings_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.AdjustLanguageSettings " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub RestoreSettings()
        '<EhHeader>
        On Error GoTo RestoreSettings_Err

        '</EhHeader>
100     If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True Then
102         chkUseDeferredDeliveryTime.Value = 1
        Else
104         chkUseDeferredDeliveryTime.Value = 0
        End If

106     If gtDeferredDeliveryTimeSettings.bSingleSMS = True Then
108         optSingleSMS.Value = True
        Else
110         optSingleSMS.Value = False
        End If

112     If gtDeferredDeliveryTimeSettings.bPeriodicSMS = True Then
114         optPeriodicSMS.Value = True
        Else
116         optPeriodicSMS.Value = False
        End If

118     cboTimeZone.ListIndex = gtDeferredDeliveryTimeSettings.nTimeZone
120     txtDeferredDeliveryTime.Text = gtDeferredDeliveryTimeSettings.sDeliveryDate
122     txtStartingDate.Text = gtDeferredDeliveryTimeSettings.sStartingDate
124     txtNumberOfMessages.Text = gtDeferredDeliveryTimeSettings.sNumberOfMessages
126     txtWaitingTime.Text = gtDeferredDeliveryTimeSettings.sWaitingTime
128     cboWaitingPeriod.ListIndex = gtDeferredDeliveryTimeSettings.nWaitingPeriod

        '<EhFooter>
        Exit Sub

RestoreSettings_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.RestoreSettings " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub SaveSettings()
        '<EhHeader>
        On Error GoTo SaveSettings_Err

        '</EhHeader>
100     If chkUseDeferredDeliveryTime.Value = 1 Then
102         gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True
        Else
104         gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = False
        End If

106     If optSingleSMS.Value = True Then
108         gtDeferredDeliveryTimeSettings.bSingleSMS = True
        Else
110         gtDeferredDeliveryTimeSettings.bSingleSMS = False
        End If

112     If optPeriodicSMS.Value = True Then
114         gtDeferredDeliveryTimeSettings.bPeriodicSMS = True
        Else
116         gtDeferredDeliveryTimeSettings.bPeriodicSMS = False
        End If

118     gtDeferredDeliveryTimeSettings.nTimeZone = cboTimeZone.ListIndex
120     gtDeferredDeliveryTimeSettings.sDeliveryDate = txtDeferredDeliveryTime.Text
122     gtDeferredDeliveryTimeSettings.sStartingDate = txtStartingDate.Text
124     gtDeferredDeliveryTimeSettings.sNumberOfMessages = txtNumberOfMessages.Text
126     gtDeferredDeliveryTimeSettings.sWaitingTime = txtWaitingTime.Text
128     gtDeferredDeliveryTimeSettings.nWaitingPeriod = cboWaitingPeriod.ListIndex

        '<EhFooter>
        Exit Sub

SaveSettings_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.SaveSettings " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub UpdateControlAppearance()
        '<EhHeader>
        On Error GoTo UpdateControlAppearance_Err

        '</EhHeader>
100     If chkUseDeferredDeliveryTime.Value = 0 Then
102         optSingleSMS.Enabled = False
104         optPeriodicSMS.Enabled = False
106         cboTimeZone.Enabled = False
108         txtDeferredDeliveryTime.Enabled = False
110         txtStartingDate.Enabled = False
112         txtNumberOfMessages.Enabled = False
114         txtWaitingTime.Enabled = False
116         cboWaitingPeriod.Enabled = False
118         lstSchedulingPreview.Enabled = False
        Else
120         lstSchedulingPreview.Enabled = True
122         optSingleSMS.Enabled = True
124         optPeriodicSMS.Enabled = True
126         cboTimeZone.Enabled = True
  
128         If optSingleSMS.Value = True Then
130             txtDeferredDeliveryTime.Enabled = True
132             txtStartingDate.Enabled = False
134             txtNumberOfMessages.Enabled = False
136             txtWaitingTime.Enabled = False
138             cboWaitingPeriod.Enabled = False
            Else
140             txtDeferredDeliveryTime.Enabled = False
142             txtStartingDate.Enabled = True
144             txtNumberOfMessages.Enabled = True
146             txtWaitingTime.Enabled = True
148             cboWaitingPeriod.Enabled = True
            End If
        End If

        '<EhFooter>
        Exit Sub

UpdateControlAppearance_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.UpdateControlAppearance " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub UpdateSchedulingPreview()
        '<EhHeader>
        On Error GoTo UpdateSchedulingPreview_Err
        '</EhHeader>
        On Error GoTo ErrorTrap
        Dim i As Integer
        Dim dWork As Date

100     If chkUseDeferredDeliveryTime.Value = 0 Then
            'Do nothing
        Else
102         lstSchedulingPreview.Clear

104         If optSingleSMS.Value = True Then
106             lstSchedulingPreview.AddItem LoadLanguageSpecificString(gnLanguage, 193) & " " & LoadLanguageSpecificString(gnLanguage, 194) & " " & txtDeferredDeliveryTime.Text
            Else

108             If IsDate(txtStartingDate.Text) Then
110                 dWork = CVDate(txtStartingDate.Text)

112                 For i = 1 To Val(txtNumberOfMessages.Text)
114                     lstSchedulingPreview.AddItem LoadLanguageSpecificString(gnLanguage, 193) & Str$(i) & " " & LoadLanguageSpecificString(gnLanguage, 194) & " " & dWork

116                     Select Case cboWaitingPeriod.ListIndex
        
                            Case 0 'Seconds
118                             dWork = DateAdd("s", txtWaitingTime.Text, dWork)
        
120                         Case 1 'Minutes
122                             dWork = DateAdd("n", txtWaitingTime.Text, dWork)
        
124                         Case 2 'Hours
126                             dWork = DateAdd("h", txtWaitingTime.Text, dWork)
        
128                         Case 3 'Days
130                             dWork = DateAdd("d", txtWaitingTime.Text, dWork)
        
132                         Case 4 'Weeks
134                             dWork = DateAdd("ww", txtWaitingTime.Text, dWork)
        
136                         Case 5 'Months
138                             dWork = DateAdd("m", txtWaitingTime.Text, dWork)
          
140                         Case Else
                                'Unexecpted, do nothing, probably an application bug
          
                        End Select

                    Next

                End If
            End If
        End If

        Exit Sub
ErrorTrap:
142     lstSchedulingPreview.Clear
144     lstSchedulingPreview.AddItem LoadLanguageSpecificString(gnLanguage, 195)
        '<EhFooter>
        Exit Sub

UpdateSchedulingPreview_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.UpdateSchedulingPreview " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cboWaitingPeriod_Change()
        '<EhHeader>
        On Error GoTo cboWaitingPeriod_Change_Err
        '</EhHeader>
100     UpdateSchedulingPreview
        '<EhFooter>
        Exit Sub

cboWaitingPeriod_Change_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.cboWaitingPeriod_Change " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cboWaitingPeriod_Click()
        '<EhHeader>
        On Error GoTo cboWaitingPeriod_Click_Err
        '</EhHeader>
100     UpdateSchedulingPreview
        '<EhFooter>
        Exit Sub

cboWaitingPeriod_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.cboWaitingPeriod_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkUseDeferredDeliveryTime_Click()
        '<EhHeader>
        On Error GoTo chkUseDeferredDeliveryTime_Click_Err
        '</EhHeader>
100     UpdateControlAppearance
        '<EhFooter>
        Exit Sub

chkUseDeferredDeliveryTime_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.chkUseDeferredDeliveryTime_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.cmdCancel_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOk_Click()
        '<EhHeader>
        On Error GoTo cmdOk_Click_Err
        '</EhHeader>
100     SaveSettings
102     UpdateJobList gnLanguage
104     Unload Me
        '<EhFooter>
        Exit Sub

cmdOk_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.cmdOk_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        Dim nDefaultTimeZone As Integer

100     CenterForm Me
102     AdjustFontControls Me

104     AdjustLanguageSettings gnLanguage

106     cboWaitingPeriod.ListIndex = 3

108     VersionSpecificAction 36, nDefaultTimeZone

110     cboTimeZone.ListIndex = nDefaultTimeZone

112     RestoreSettings

114     UpdateControlAppearance
116     UpdateSchedulingPreview

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub optPeriodicSMS_Click()
        '<EhHeader>
        On Error GoTo optPeriodicSMS_Click_Err
        '</EhHeader>
100     UpdateControlAppearance
102     UpdateSchedulingPreview
        '<EhFooter>
        Exit Sub

optPeriodicSMS_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.optPeriodicSMS_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub optSingleSMS_Click()
        '<EhHeader>
        On Error GoTo optSingleSMS_Click_Err
        '</EhHeader>
100     UpdateControlAppearance
102     UpdateSchedulingPreview
        '<EhFooter>
        Exit Sub

optSingleSMS_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.optSingleSMS_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtDeferredDeliveryTime_Change()
        '<EhHeader>
        On Error GoTo txtDeferredDeliveryTime_Change_Err
        '</EhHeader>
100     UpdateSchedulingPreview
        '<EhFooter>
        Exit Sub

txtDeferredDeliveryTime_Change_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.txtDeferredDeliveryTime_Change " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtNumberOfMessages_Change()
        '<EhHeader>
        On Error GoTo txtNumberOfMessages_Change_Err
        '</EhHeader>
100     UpdateSchedulingPreview
        '<EhFooter>
        Exit Sub

txtNumberOfMessages_Change_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.txtNumberOfMessages_Change " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtStartingDate_Change()
        '<EhHeader>
        On Error GoTo txtStartingDate_Change_Err
        '</EhHeader>
100     UpdateSchedulingPreview
        '<EhFooter>
        Exit Sub

txtStartingDate_Change_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.txtStartingDate_Change " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtWaitingTime_Change()
        '<EhHeader>
        On Error GoTo txtWaitingTime_Change_Err
        '</EhHeader>
100     UpdateSchedulingPreview
        '<EhFooter>
        Exit Sub

txtWaitingTime_Change_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_SMS_Messenger.frmDeferredDeliveryTime.txtWaitingTime_Change " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

