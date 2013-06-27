VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDBAttributeSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database options"
   ClientHeight    =   4155
   ClientLeft      =   5025
   ClientTop       =   1980
   ClientWidth     =   6795
   Icon            =   "frmDBAttributeSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDatabaseOptions 
      Caption         =   "Visible Database fields"
      Height          =   3375
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   6495
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   16
         Left            =   3360
         TabIndex        =   15
         Top             =   2880
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   15
         Left            =   3360
         TabIndex        =   14
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   14
         Left            =   3360
         TabIndex        =   13
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   13
         Left            =   3360
         TabIndex        =   12
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   12
         Left            =   3360
         TabIndex        =   11
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   11
         Left            =   3360
         TabIndex        =   10
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   10
         Left            =   3360
         TabIndex        =   9
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   9
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox chkAttribute 
         Caption         =   "chkAttribute"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdSaveSettings 
      Caption         =   "Save Settings"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3600
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog cmdDlgOpen 
      Left            =   4080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDBAttributeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AdjustLanguageSettings(nLanguage As Integer)
        '<EhHeader>
        On Error GoTo AdjustLanguageSettings_Err
        '</EhHeader>
        Dim sTemp As String
        Dim sApplicationPlaceHolder As String

100     Me.Caption = LoadLanguageSpecificString(nLanguage, 622)
102     cmdSaveSettings.Caption = LoadLanguageSpecificString(nLanguage, 623)
104     cmdClose.Caption = LoadLanguageSpecificString(nLanguage, 624)
        '<EhFooter>
        Exit Sub

AdjustLanguageSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDBAttributeSettings.AdjustLanguageSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub AdjustLanguageSettingsJobLog(nLanguage As Integer)
        '<EhHeader>
        On Error GoTo AdjustLanguageSettingsJobLog_Err
        '</EhHeader>
        Dim i As Integer
100     Me.Tag = "2"
102     chkAttribute(1).Caption = LoadLanguageSpecificString(nLanguage, 560)
104     chkAttribute(2).Caption = GetJobRemarksFieldFromDatabase(nLanguage, 1)
106     chkAttribute(3).Caption = GetJobRemarksFieldFromDatabase(nLanguage, 2)
108     chkAttribute(4).Caption = GetJobRemarksFieldFromDatabase(nLanguage, 3)
110     chkAttribute(5).Caption = GetJobRemarksFieldFromDatabase(nLanguage, 4)
112     chkAttribute(6).Caption = GetJobRemarksFieldFromDatabase(nLanguage, 5)
114     chkAttribute(7).Caption = GetJobRemarksFieldFromDatabase(nLanguage, 6)
116     chkAttribute(8).Caption = LoadLanguageSpecificString(nLanguage, 581)
118     chkAttribute(9).Caption = LoadLanguageSpecificString(nLanguage, 571)
120     chkAttribute(10).Caption = LoadLanguageSpecificString(nLanguage, 570)
122     chkAttribute(11).Caption = LoadLanguageSpecificString(nLanguage, 568)
124     chkAttribute(12).Caption = LoadLanguageSpecificString(nLanguage, 573)
126     chkAttribute(13).Caption = LoadLanguageSpecificString(nLanguage, 574)
128     chkAttribute(14).Caption = LoadLanguageSpecificString(nLanguage, 575)
130     chkAttribute(15).Caption = LoadLanguageSpecificString(nLanguage, 576)
132     chkAttribute(16).Caption = LoadLanguageSpecificString(nLanguage, 569)

134     For i = 1 To 16

136         If gtDBLogFieldsSettings.bJobLog(i) = True Then
138             chkAttribute(i).Value = 1
            Else
140             chkAttribute(i).Value = 0
            End If

        Next

142     fraDatabaseOptions.Caption = LoadLanguageSpecificString(nLanguage, 621)
        '<EhFooter>
        Exit Sub

AdjustLanguageSettingsJobLog_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDBAttributeSettings.AdjustLanguageSettingsJobLog " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub AdjustLanguageSettingsSendLog(nLanguage As Integer)
        '<EhHeader>
        On Error GoTo AdjustLanguageSettingsSendLog_Err
        '</EhHeader>
        Dim i As Integer
100     chkAttribute(1).Caption = LoadLanguageSpecificString(nLanguage, 614)
102     chkAttribute(2).Caption = LoadLanguageSpecificString(nLanguage, 615)
104     chkAttribute(3).Caption = LoadLanguageSpecificString(nLanguage, 600)
106     chkAttribute(4).Caption = LoadLanguageSpecificString(nLanguage, 601)
108     chkAttribute(5).Caption = LoadLanguageSpecificString(nLanguage, 602)
110     chkAttribute(6).Caption = LoadLanguageSpecificString(nLanguage, 603)
112     chkAttribute(7).Caption = LoadLanguageSpecificString(nLanguage, 604)
114     chkAttribute(8).Caption = LoadLanguageSpecificString(nLanguage, 605)
116     chkAttribute(9).Caption = LoadLanguageSpecificString(nLanguage, 606)
118     chkAttribute(10).Caption = LoadLanguageSpecificString(nLanguage, 607)
120     chkAttribute(11).Caption = LoadLanguageSpecificString(nLanguage, 608)
122     chkAttribute(12).Caption = LoadLanguageSpecificString(nLanguage, 609)
124     chkAttribute(13).Caption = LoadLanguageSpecificString(nLanguage, 610)
126     chkAttribute(14).Caption = LoadLanguageSpecificString(nLanguage, 611)
128     chkAttribute(15).Caption = LoadLanguageSpecificString(nLanguage, 612)
130     chkAttribute(16).Caption = LoadLanguageSpecificString(nLanguage, 613)

132     For i = 1 To 16

134         If gtDBLogFieldsSettings.bSendLog(i) = True Then
136             chkAttribute(i).Value = 1
            Else
138             chkAttribute(i).Value = 0
            End If

        Next

140     Me.Tag = "1"
142     fraDatabaseOptions.Caption = LoadLanguageSpecificString(nLanguage, 620)
        '<EhFooter>
        Exit Sub

AdjustLanguageSettingsSendLog_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDBAttributeSettings.AdjustLanguageSettingsSendLog " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub SaveLogSettings()
        '<EhHeader>
        On Error GoTo SaveLogSettings_Err
        '</EhHeader>
        Dim sRelevantSetting As String
        Dim i As Integer

100     For i = 1 To 16

102         Select Case Me.Tag

                Case "1"  'Sendlog Settings
104                 sRelevantSetting = "SendLogFieldEnabled" & Right("00" & Trim(Str$(i)), 2)
    
106             Case "2" 'Joglog Settings
108                 sRelevantSetting = "JobLogFieldEnabled" & Right("00" & Trim(Str$(i)), 2)
  
110             Case Else
                    'Do nothing
    
            End Select
  
112         If chkAttribute(i).Value = 1 Then
114             PutSettingIntoDataBase sRelevantSetting, True
            Else
116             PutSettingIntoDataBase sRelevantSetting, False
            End If

        Next

        '<EhFooter>
        Exit Sub

SaveLogSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDBAttributeSettings.SaveLogSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdClose_Click()
        '<EhHeader>
        On Error GoTo cmdClose_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

cmdClose_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDBAttributeSettings.cmdClose_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSaveSettings_Click()
        '<EhHeader>
        On Error GoTo cmdSaveSettings_Click_Err
        '</EhHeader>
100     SaveLogSettings
102     RestoreLogSettings

104     Select Case Me.Tag

            Case "1"
106             InitSendLogColWidthSettings
  
108         Case "2"
110             InitJobLogColWidthSettings
  
        End Select

112     Unload Me
        '<EhFooter>
        Exit Sub

cmdSaveSettings_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDBAttributeSettings.cmdSaveSettings_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     CenterForm Me
102     AdjustFontControls Me
104     AdjustLanguageSettings gnLanguage
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmDBAttributeSettings.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

