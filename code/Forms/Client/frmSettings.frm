VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Settings"
   ClientHeight    =   6000
   ClientLeft      =   3990
   ClientTop       =   3705
   ClientWidth     =   7350
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabSettings 
      Height          =   5295
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "Usersettings"
      TabPicture(0)   =   "frmSettings.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblUserkey"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPassword"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblOriginator"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkSaveMessagesInSendJournal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtUserkey"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPassword"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtOriginator"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOriginatorCheck"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Language"
      TabPicture(1)   =   "frmSettings.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optLanguage(1)"
      Tab(1).Control(1)=   "optLanguage(2)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Individual Font"
      TabPicture(2)   =   "frmSettings.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdSelectFont"
      Tab(2).Control(1)=   "chkUseIndividualFont"
      Tab(2).Control(2)=   "lblFontPreview"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Active SMS Types"
      TabPicture(3)   =   "frmSettings.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkSMSTypeEnabled(0)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Userdefined Database Fields"
      TabPicture(4)   =   "frmSettings.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraJobRemarks"
      Tab(4).Control(1)=   "fraPhonebook"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Validity Period"
      TabPicture(5)   =   "frmSettings.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraSingleSMS"
      Tab(5).Control(1)=   "fraPeriodicSMS"
      Tab(5).ControlCount=   2
      Begin VB.CommandButton cmdOriginatorCheck 
         Caption         =   "..."
         Height          =   255
         Left            =   3600
         TabIndex        =   53
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame fraPeriodicSMS 
         Caption         =   "Periodic Sending"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   36
         Top             =   2520
         Width           =   6735
         Begin VB.CheckBox chkPeriodicSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes 
            Caption         =   "Use userdefined Validity Period only for specific SMS Types"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Visible         =   0   'False
            Width           =   6255
         End
         Begin VB.OptionButton optPeriodicSMSUseSingleshotAsLifeTime 
            Caption         =   "Use Validity Period ""Singleshot"""
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   6015
         End
         Begin VB.OptionButton optPeriodicSMSUseSpecificSettingsAsLifeTime 
            Caption         =   "Use specific Validity Period"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   6495
         End
         Begin VB.OptionButton optPeriodicSMSUseWaitingTimeAsLifeTime 
            Caption         =   "Use ""Waiting time between messages"" as Validity Period"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Value           =   -1  'True
            Width           =   6495
         End
         Begin VB.CheckBox chkPeriodicSMSUseUserDefinedLifeTime 
            Caption         =   "Use userdefined Settings when sending periodic SMS"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   6015
         End
         Begin VB.TextBox txtPeriodicSMSSpecificSettingLifeTime 
            Height          =   285
            Left            =   2640
            TabIndex        =   33
            Text            =   "24"
            Top             =   1320
            Width           =   975
         End
         Begin VB.ComboBox cboPeriodicSMSSpecificSettingLifeTimeUnit 
            Height          =   315
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblPeriodicSMSSpecificLifeTime 
            Caption         =   "Validity Period"
            Height          =   255
            Left            =   360
            TabIndex        =   32
            Top             =   1320
            Width           =   2895
         End
      End
      Begin VB.Frame fraSingleSMS 
         Caption         =   "Single Sending"
         Height          =   1755
         Left            =   -74880
         TabIndex        =   27
         Top             =   720
         Width           =   6735
         Begin VB.CheckBox chkSingleSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes 
            Caption         =   "Use userdefined Validity Period only for specific SMS Types"
            Height          =   250
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Visible         =   0   'False
            Width           =   6495
         End
         Begin VB.OptionButton optSingleSMSUseSingleshotAsLifeTime 
            Caption         =   "Use Validity Period ""Singleshot"""
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1335
            Width           =   6015
         End
         Begin VB.OptionButton optSingleSMSUseSpecificSettingsAsLifeTime 
            Caption         =   "Use specific Validity Period"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   585
            Value           =   -1  'True
            Width           =   6495
         End
         Begin VB.TextBox txtSingleSMSSpecificSettingLifeTime 
            Height          =   285
            Left            =   2640
            TabIndex        =   24
            Text            =   "24"
            Top             =   945
            Width           =   975
         End
         Begin VB.ComboBox cboSingleSMSSpecificSettingLifeTimeUnit 
            Height          =   315
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   945
            Width           =   1095
         End
         Begin VB.CheckBox chkSingleSMSUseUserDefinedLifeTime 
            Caption         =   "Use userdefined Settings when sending single SMS"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   6375
         End
         Begin VB.Label lblSingleSMSSpecificLifeTime 
            Caption         =   "Validity Period"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   945
            Width           =   2895
         End
      End
      Begin VB.Frame fraPhonebook 
         Caption         =   "Phonebook / Personalized SMS"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   6
         Top             =   720
         Width           =   4695
         Begin VB.TextBox txtPhonebookVariableField 
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   100
            TabIndex        =   1
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox txtPhonebookVariableField 
            Height          =   285
            Index           =   2
            Left            =   2400
            MaxLength       =   100
            TabIndex        =   3
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtPhonebookVariableField 
            Height          =   285
            Index           =   3
            Left            =   2400
            MaxLength       =   100
            TabIndex        =   5
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label lblPhonebookVariableField 
            Caption         =   "lblPhonebookVariableField"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblPhonebookVariableField 
            Caption         =   "lblPhonebookVariableField"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label lblPhonebookVariableField 
            Caption         =   "lblPhonebookVariableField"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   4
            Top             =   960
            Width           =   2175
         End
      End
      Begin VB.Frame fraJobRemarks 
         Caption         =   "Jobinfo"
         Height          =   2415
         Left            =   -74880
         TabIndex        =   19
         Top             =   2160
         Width           =   4695
         Begin VB.TextBox txtJobRemarks 
            Height          =   285
            Index           =   6
            Left            =   2400
            MaxLength       =   100
            TabIndex        =   18
            Top             =   2040
            Width           =   2175
         End
         Begin VB.TextBox txtJobRemarks 
            Height          =   285
            Index           =   5
            Left            =   2400
            MaxLength       =   100
            TabIndex        =   16
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox txtJobRemarks 
            Height          =   285
            Index           =   4
            Left            =   2400
            MaxLength       =   100
            TabIndex        =   14
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox txtJobRemarks 
            Height          =   285
            Index           =   3
            Left            =   2400
            MaxLength       =   100
            TabIndex        =   12
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtJobRemarks 
            Height          =   285
            Index           =   2
            Left            =   2400
            MaxLength       =   100
            TabIndex        =   10
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtJobRemarks 
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   100
            TabIndex        =   8
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   17
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label lblJobRemarks 
            Caption         =   "lblJobRemarks"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CheckBox chkSMSTypeEnabled 
         Caption         =   "Text SMS"
         Height          =   375
         Index           =   0
         Left            =   -74880
         TabIndex        =   0
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdSelectFont 
         Caption         =   "Select Font..."
         Height          =   375
         Left            =   -74880
         TabIndex        =   49
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox chkUseIndividualFont 
         Caption         =   "Use individual Font"
         Height          =   255
         Left            =   -74880
         TabIndex        =   48
         Top             =   1440
         Width           =   3615
      End
      Begin VB.OptionButton optLanguage 
         Caption         =   "German"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   46
         Top             =   1080
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.OptionButton optLanguage 
         Caption         =   "English"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   45
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox txtOriginator 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   42
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   40
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtUserkey 
         Height          =   285
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   38
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkSaveMessagesInSendJournal 
         Caption         =   "Save messages in work journal"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1920
         Value           =   1  'Checked
         Width           =   4695
      End
      Begin VB.Label lblFontPreview 
         Caption         =   "Current Font"
         Height          =   495
         Left            =   -74880
         TabIndex        =   47
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label lblOriginator 
         Caption         =   "Originator:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblUserkey 
         Caption         =   "Userkey:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   52
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Save Settings"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Top             =   5520
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog cmDlgFonts 
      Left            =   4680
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mnLanguage As Integer

Public Sub AdjustLanguageSettings(nLanguage As Integer)
        '<EhHeader>
        On Error GoTo AdjustLanguageSettings_Err
        '</EhHeader>
        Dim i As Integer
100     mnLanguage = nLanguage
102     Me.caption = LoadLanguageSpecificString(nLanguage, 51)
104     tabSettings.TabCaption(0) = LoadLanguageSpecificString(nLanguage, 52)
106     tabSettings.TabCaption(1) = LoadLanguageSpecificString(nLanguage, 53)
108     tabSettings.TabCaption(2) = LoadLanguageSpecificString(nLanguage, 58)
110     tabSettings.TabCaption(3) = LoadLanguageSpecificString(nLanguage, 61)
112     tabSettings.TabCaption(4) = LoadLanguageSpecificString(nLanguage, 73)
114     tabSettings.TabCaption(5) = LoadLanguageSpecificString(nLanguage, 681)

116     lblUserkey.caption = LoadLanguageSpecificString(nLanguage, 54)
118     lblPassword.caption = LoadLanguageSpecificString(nLanguage, 55)
       ' lblOriginator.Caption = LoadLanguageSpecificString(nLanguage, 56)
       ' chkSaveMessagesInSendJournal.Caption = LoadLanguageSpecificString(nLanguage, 57)

120     optLanguage(1).caption = LoadLanguageSpecificString(nLanguage, 3)
        'optLanguage(2).Caption = LoadLanguageSpecificString(nLanguage, 4)

122     chkUseIndividualFont.caption = LoadLanguageSpecificString(nLanguage, 59)
124     cmdSelectFont.caption = LoadLanguageSpecificString(nLanguage, 60)

126     chkSMSTypeEnabled(0).caption = LoadLanguageSpecificString(nLanguage, 62)
128     chkSMSTypeEnabled(1).caption = LoadLanguageSpecificString(nLanguage, 63)
130     chkSMSTypeEnabled(2).caption = LoadLanguageSpecificString(nLanguage, 64)
132     chkSMSTypeEnabled(3).caption = LoadLanguageSpecificString(nLanguage, 65)
134     chkSMSTypeEnabled(4).caption = LoadLanguageSpecificString(nLanguage, 66)
136     chkSMSTypeEnabled(5).caption = LoadLanguageSpecificString(nLanguage, 67)
138     chkSMSTypeEnabled(6).caption = LoadLanguageSpecificString(nLanguage, 480)
140     chkSMSTypeEnabled(7).caption = LoadLanguageSpecificString(nLanguage, 478)
142     chkSMSTypeEnabled(8).caption = LoadLanguageSpecificString(nLanguage, 68)
144     VersionSpecificAction 7, nLanguage

146     cmdOk.caption = LoadLanguageSpecificString(nLanguage, 71)
148     cmdCancel.caption = LoadLanguageSpecificString(nLanguage, 72)

150     fraJobRemarks.caption = LoadLanguageSpecificString(nLanguage, 74)

152     For i = 1 To gcnNumberOfJobRemarkFields
154         lblJobRemarks(i).caption = LoadLanguageSpecificString(nLanguage, 560 + i)
        Next

156     For i = 1 To gcnNumberOfJobRemarkFields

158         If GetSettingFromDatabase("JobRemarksField0" & Trim(Str$(i))) = "DEFAULT" Then
160             txtJobRemarks(i).Text = GetJobRemarksFieldFromDatabase(nLanguage, i)
            Else
162             txtJobRemarks(i).Text = GetSettingFromDatabase("JobRemarksField0" & Trim(Str$(i)))
            End If

        Next

164     For i = 1 To 3

166         If GetSettingFromDatabase("PhonebookVariableField0" & Trim(Str$(i))) = "DEFAULT" Then
168             txtPhonebookVariableField(i).Text = GetPhonebookVariableFieldFromDatabase(nLanguage, i)
            Else
170             txtPhonebookVariableField(i).Text = GetSettingFromDatabase("PhonebookVariableField0" & Trim(Str$(i)))
            End If

        Next

       ' fraPhonebook.Caption = LoadLanguageSpecificString(nLanguage, 75)
172     lblPhonebookVariableField(1).caption = LoadLanguageSpecificString(nLanguage, 76)
174     lblPhonebookVariableField(2).caption = LoadLanguageSpecificString(nLanguage, 77)
176     lblPhonebookVariableField(3).caption = LoadLanguageSpecificString(nLanguage, 78)

178     fraSingleSMS.caption = LoadLanguageSpecificString(nLanguage, 682)
180     chkSingleSMSUseUserDefinedLifeTime.caption = LoadLanguageSpecificString(nLanguage, 683)
182     optSingleSMSUseSpecificSettingsAsLifeTime.caption = LoadLanguageSpecificString(nLanguage, 685)
184     lblSingleSMSSpecificLifeTime.caption = LoadLanguageSpecificString(nLanguage, 686)
186     cboSingleSMSSpecificSettingLifeTimeUnit.Clear
188     cboSingleSMSSpecificSettingLifeTimeUnit.AddItem LoadLanguageSpecificString(nLanguage, 188)
190     cboSingleSMSSpecificSettingLifeTimeUnit.AddItem LoadLanguageSpecificString(nLanguage, 189)
192     cboSingleSMSSpecificSettingLifeTimeUnit.ListIndex = 1
194     optSingleSMSUseSingleshotAsLifeTime.caption = LoadLanguageSpecificString(nLanguage, 687)
196     chkSingleSMSUseUserDefinedLifeTime_Click

198     fraPeriodicSMS.caption = LoadLanguageSpecificString(nLanguage, 691)
200     chkPeriodicSMSUseUserDefinedLifeTime.caption = LoadLanguageSpecificString(nLanguage, 692)
202     optPeriodicSMSUseWaitingTimeAsLifeTime.caption = LoadLanguageSpecificString(nLanguage, 694)
204     optPeriodicSMSUseSpecificSettingsAsLifeTime.caption = LoadLanguageSpecificString(nLanguage, 695)
206     lblPeriodicSMSSpecificLifeTime.caption = LoadLanguageSpecificString(nLanguage, 696)
208     cboPeriodicSMSSpecificSettingLifeTimeUnit.Clear
210     cboPeriodicSMSSpecificSettingLifeTimeUnit.AddItem LoadLanguageSpecificString(nLanguage, 188)
212     cboPeriodicSMSSpecificSettingLifeTimeUnit.AddItem LoadLanguageSpecificString(nLanguage, 189)
214     cboPeriodicSMSSpecificSettingLifeTimeUnit.ListIndex = 1
216     chkPeriodicSMSUseUserDefinedLifeTime_Click

218     optPeriodicSMSUseSingleshotAsLifeTime.caption = LoadLanguageSpecificString(nLanguage, 697)
        '<EhFooter>
        Exit Sub

AdjustLanguageSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.AdjustLanguageSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function CheckScreenInput() As Boolean
        'Check, if each description is different
        '<EhHeader>
        On Error GoTo CheckScreenInput_Err
        '</EhHeader>
        Dim sFieldTest() As String
        Dim i As Integer
        Dim k As Integer
        Dim bIdenticalEntryFound As Boolean
        Dim bEmptyEntryFound As Boolean

100     For i = txtPhonebookVariableField.LBound To txtPhonebookVariableField.UBound
102         For k = txtPhonebookVariableField.LBound To txtPhonebookVariableField.UBound

104             If i <> k Then
106                 If Trim(UCase(SQLValidCharsEncode(txtPhonebookVariableField(i).Text))) <> Trim(UCase(SQLValidCharsEncode(txtPhonebookVariableField(k).Text))) Then
                        'Ok
                    Else
108                     bIdenticalEntryFound = True
                    End If
                End If

            Next

110         If Trim(UCase(SQLValidCharsEncode(txtPhonebookVariableField(i).Text))) = "" Then
112             bEmptyEntryFound = True
            End If

        Next

114     For i = txtJobRemarks.LBound To txtJobRemarks.UBound
116         For k = txtJobRemarks.LBound To txtJobRemarks.UBound

118             If i <> k Then
120                 If Trim(UCase(SQLValidCharsEncode(txtJobRemarks(i).Text))) <> Trim(UCase(SQLValidCharsEncode(txtJobRemarks(k).Text))) Then
                        'Ok
                    Else
122                     bIdenticalEntryFound = True
                    End If
                End If

            Next

124         If Trim(UCase(SQLValidCharsEncode(txtJobRemarks(i).Text))) = "" Then
126             bEmptyEntryFound = True
            End If

        Next

128     If bIdenticalEntryFound = True Then
130         MsgBox LoadLanguageSpecificString(gnLanguage, 79), vbCritical, gsApplicationName
132         CheckScreenInput = False
            Exit Function
        Else
            'Do nothing, Continue
        End If

134     If bEmptyEntryFound = True Then
136         MsgBox LoadLanguageSpecificString(gnLanguage, 80), vbCritical, gsApplicationName
138         CheckScreenInput = False
            Exit Function
        Else
            'Do nothing, Continue
        End If

140     CheckScreenInput = True
        '<EhFooter>
        Exit Function

CheckScreenInput_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.CheckScreenInput " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub SaveGeneralSettings()
        '<EhHeader>
        On Error GoTo SaveGeneralSettings_Err
        '</EhHeader>
        Dim sVersionSpecific1 As String
        Dim sVersionSpecific2 As String
        Dim sVersionSpecific3 As String
        Dim nLanguage As Integer
        Dim tValidityPeriodOrig As ValidityPeriod
        Dim tValidityPeriodTrimmed As ValidityPeriod
        Dim bValidityPeriodHasBeenTrimmed As Boolean

        Dim i As Integer

100     VersionSpecificAction 17, , , sVersionSpecific1
102     VersionSpecificAction 18, , , sVersionSpecific2
104     VersionSpecificAction 19, , , sVersionSpecific3

106     PutSettingIntoDataBase "Userkey", txtUserkey.Text
108     PutSettingIntoDataBase "Password", txtPassword.Text
110     PutSettingIntoDataBase "Originator", txtOriginator.Text

112     If chkSaveMessagesInSendJournal.Value = 1 Then
114         PutSettingIntoDataBase "SaveMessagesInSendJournal", True
        Else
116         PutSettingIntoDataBase "SaveMessagesInSendJournal", False
        End If

118     Select Case True

            Case optLanguage(1).Value = True
120             PutSettingIntoDataBase "Language", 1
122             nLanguage = 1
  
124         Case optLanguage(2).Value = True
126             PutSettingIntoDataBase "Language", 2
128             nLanguage = 2
  
        End Select

130     If chkSMSTypeEnabled(0).Value = 1 Then
132         PutSettingIntoDataBase "TextSMSEnabled", True
        Else
134         PutSettingIntoDataBase "TextSMSEnabled", False
        End If

136     If chkSMSTypeEnabled(1).Value = 1 Then
138         PutSettingIntoDataBase "OperatorLogoEnabled", True
        Else
140         PutSettingIntoDataBase "OperatorLogoEnabled", False
        End If

142     If chkSMSTypeEnabled(2).Value = 1 Then
144         PutSettingIntoDataBase "GroupLogoEnabled", True
        Else
146         PutSettingIntoDataBase "GroupLogoEnabled", False
        End If

148     If chkSMSTypeEnabled(3).Value = 1 Then
150         PutSettingIntoDataBase "RingtoneEnabled", True
        Else
152         PutSettingIntoDataBase "RingtoneEnabled", False
        End If

154     If chkSMSTypeEnabled(4).Value = 1 Then
156         PutSettingIntoDataBase "PictureMessageEnabled", True
        Else
158         PutSettingIntoDataBase "PictureMessageEnabled", False
        End If

160     If chkSMSTypeEnabled(5).Value = 1 Then
162         PutSettingIntoDataBase "VCardEnabled", True
        Else
164         PutSettingIntoDataBase "VCardEnabled", False
        End If

166     If chkSMSTypeEnabled(6).Value = 1 Then
168         PutSettingIntoDataBase "UnicodeEnabled", True
        Else
170         PutSettingIntoDataBase "UnicodeEnabled", False
        End If

172     If chkSMSTypeEnabled(7).Value = 1 Then
174         PutSettingIntoDataBase "WAPPushSMSEnabled", True
        Else
176         PutSettingIntoDataBase "WAPPushSMSEnabled", False
        End If

178     If chkSMSTypeEnabled(8).Value = 1 Then
180         PutSettingIntoDataBase "BinaryDataEnabled", True
        Else
182         PutSettingIntoDataBase "BinaryDataEnabled", False
        End If

184     If ControlIsLoaded(chkSMSTypeEnabled(9)) Then
186         If chkSMSTypeEnabled(9).Value = 1 Then
188             PutSettingIntoDataBase sVersionSpecific1, True
            Else
190             PutSettingIntoDataBase sVersionSpecific1, False
            End If
        End If

192     If ControlIsLoaded(chkSMSTypeEnabled(10)) Then
194         If chkSMSTypeEnabled(10).Value = 1 Then
196             PutSettingIntoDataBase sVersionSpecific2, True
            Else
198             PutSettingIntoDataBase sVersionSpecific2, False
            End If
        End If

200     If ControlIsLoaded(chkSMSTypeEnabled(11)) Then
202         If chkSMSTypeEnabled(11).Value = 1 Then
204             PutSettingIntoDataBase sVersionSpecific3, True
            Else
206             PutSettingIntoDataBase sVersionSpecific3, False
            End If
        End If

208     If chkUseIndividualFont.Value = 1 Then
210         PutSettingIntoDataBase "SpecificFontUsed", True
        Else
212         PutSettingIntoDataBase "SpecificFontUsed", False
        End If

214     For i = 1 To gcnNumberOfJobRemarkFields
216         PutJobRemarksFieldIntoDatabase nLanguage, i, SQLValidCharsEncode(txtJobRemarks(i).Text)
        Next

218     If txtPhonebookVariableField(1).Text = LoadLanguageSpecificString(nLanguage, 213) Then
220         PutSettingIntoDataBase "PhonebookVariableField01", "DEFAULT"
        Else
222         PutSettingIntoDataBase "PhonebookVariableField01", SQLValidCharsEncode(txtPhonebookVariableField(1).Text)
        End If

224     If txtPhonebookVariableField(2).Text = LoadLanguageSpecificString(nLanguage, 214) Then
226         PutSettingIntoDataBase "PhonebookVariableField02", "DEFAULT"
        Else
228         PutSettingIntoDataBase "PhonebookVariableField02", SQLValidCharsEncode(txtPhonebookVariableField(2).Text)
        End If

230     If txtPhonebookVariableField(3).Text = LoadLanguageSpecificString(nLanguage, 215) Then
232         PutSettingIntoDataBase "PhonebookVariableField03", "DEFAULT"
        Else
234         PutSettingIntoDataBase "PhonebookVariableField03", SQLValidCharsEncode(txtPhonebookVariableField(3).Text)
        End If

236     PutSettingIntoDataBase "FontName", cmDlgFonts.FontName
238     PutSettingIntoDataBase "FontSize", cmDlgFonts.FontSize
240     PutSettingIntoDataBase "FontBold", cmDlgFonts.FontBold
242     PutSettingIntoDataBase "FontItalic", cmDlgFonts.FontItalic
244     PutSettingIntoDataBase "FontStrikethru", cmDlgFonts.FontStrikethru
246     PutSettingIntoDataBase "FontUnderline", cmDlgFonts.FontUnderline

248     If chkSingleSMSUseUserDefinedLifeTime.Value = 1 Then
250         PutSettingIntoDataBase "SingleSMSUseUserDefinedLifeTime", True
        Else
252         PutSettingIntoDataBase "SingleSMSUseUserDefinedLifeTime", False
        End If

254     If chkSingleSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes.Value = 1 Then
256         PutSettingIntoDataBase "SingleSMSUseUDSettingsOnlyForSpecificSMSTypes", True
        Else
258         PutSettingIntoDataBase "SingleSMSUseUDSettingsOnlyForSpecificSMSTypes", False
        End If

260     Select Case True

            Case optSingleSMSUseSpecificSettingsAsLifeTime.Value = True
262             PutSettingIntoDataBase "SingleSMSValidityPeriodMode", ValidityPeriodMode.UseSpecificSettingsAsLifeTime
  
264         Case optSingleSMSUseSingleshotAsLifeTime.Value = True
266             PutSettingIntoDataBase "SingleSMSValidityPeriodMode", ValidityPeriodMode.UseSingleshotAsLifeTime
  
268         Case Else
                'Do nothing

        End Select

270     PutSettingIntoDataBase "SingleSMSSpecificSettingLifetime", Val(txtSingleSMSSpecificSettingLifeTime.Text)

272     tValidityPeriodOrig.nLifeTimeUnit = cboSingleSMSSpecificSettingLifeTimeUnit.ListIndex
274     tValidityPeriodOrig.lLifeTime = Val(txtSingleSMSSpecificSettingLifeTime.Text)
276     tValidityPeriodTrimmed = TrimValidityPeriodToCorrectValues(tValidityPeriodOrig)

278     If tValidityPeriodOrig.nLifeTimeUnit <> tValidityPeriodTrimmed.nLifeTimeUnit Or (tValidityPeriodOrig.lLifeTime <> tValidityPeriodTrimmed.lLifeTime) Then
280         bValidityPeriodHasBeenTrimmed = True
282         PutSettingIntoDataBase "SingleSMSSpecificSettingLifeTimeUnit", tValidityPeriodTrimmed.nLifeTimeUnit
284         PutSettingIntoDataBase "SingleSMSSpecificSettingLifetime", tValidityPeriodTrimmed.lLifeTime
        Else
286         PutSettingIntoDataBase "SingleSMSSpecificSettingLifeTimeUnit", tValidityPeriodOrig.nLifeTimeUnit
288         PutSettingIntoDataBase "SingleSMSSpecificSettingLifetime", tValidityPeriodOrig.lLifeTime
        End If

290     If chkPeriodicSMSUseUserDefinedLifeTime.Value = 1 Then
292         PutSettingIntoDataBase "PeriodicSMSUseUserDefinedLifetime", True
        Else
294         PutSettingIntoDataBase "PeriodicSMSUseUserDefinedLifetime", False
        End If

296     If chkPeriodicSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes.Value = 1 Then
298         PutSettingIntoDataBase "PeriodicSMSUseUDSettingsOnlyForSpecificSMSTypes", True
        Else
300         PutSettingIntoDataBase "PeriodicSMSUseUDSettingsOnlyForSpecificSMSTypes", False
        End If

302     Select Case True

            Case Me.optPeriodicSMSUseWaitingTimeAsLifeTime.Value = True
304             PutSettingIntoDataBase "PeriodicSMSValidityPeriodMode", ValidityPeriodMode.UseWaitingTimeAsLifeTime

306         Case optPeriodicSMSUseSpecificSettingsAsLifeTime.Value = True
308             PutSettingIntoDataBase "PeriodicSMSValidityPeriodMode", ValidityPeriodMode.UseSpecificSettingsAsLifeTime
  
310         Case optPeriodicSMSUseSingleshotAsLifeTime.Value = True
312             PutSettingIntoDataBase "PeriodicSMSValidityPeriodMode", ValidityPeriodMode.UseSingleshotAsLifeTime
  
314         Case Else
                'MsgBox "Case else"

        End Select

316     tValidityPeriodOrig.nLifeTimeUnit = cboPeriodicSMSSpecificSettingLifeTimeUnit.ListIndex
318     tValidityPeriodOrig.lLifeTime = Val(txtPeriodicSMSSpecificSettingLifeTime.Text)
320     tValidityPeriodTrimmed = TrimValidityPeriodToCorrectValues(tValidityPeriodOrig)

322     If (tValidityPeriodOrig.nLifeTimeUnit <> tValidityPeriodTrimmed.nLifeTimeUnit) Or (tValidityPeriodOrig.lLifeTime <> tValidityPeriodTrimmed.lLifeTime) Then
324         bValidityPeriodHasBeenTrimmed = True
326         PutSettingIntoDataBase "PeriodicSMSSpecificSettingLifeTimeUnit", tValidityPeriodTrimmed.nLifeTimeUnit
328         PutSettingIntoDataBase "PeriodicSMSSpecificSettingLifeTime", tValidityPeriodTrimmed.lLifeTime
        Else
330         PutSettingIntoDataBase "PeriodicSMSSpecificSettingLifeTimeUnit", tValidityPeriodOrig.nLifeTimeUnit
332         PutSettingIntoDataBase "PeriodicSMSSpecificSettingLifeTime", tValidityPeriodOrig.lLifeTime
        End If

334     If bValidityPeriodHasBeenTrimmed Then
336         MsgBox LoadLanguageSpecificString(nLanguage, 698), vbExclamation, gsApplicationName
        End If

        '<EhFooter>
        Exit Sub

SaveGeneralSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.SaveGeneralSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub SetTabVisibility(nTab As Integer, _
                     bVisible As Boolean)
        '<EhHeader>
        On Error GoTo SetTabVisibility_Err
        '</EhHeader>
        Dim i As Integer

100     Select Case nTab

            Case 0
102             lblUserkey.Visible = bVisible
104             lblPassword.Visible = bVisible
106             lblUserkey.Visible = bVisible
108             lblPassword.Visible = bVisible
110             lblOriginator.Visible = bVisible
112             chkSaveMessagesInSendJournal.Visible = bVisible
  
114         Case 1
116             optLanguage(1).Visible = bVisible
                'optLanguage(2).Visible = bVisible
  
118         Case 2
120             lblFontPreview.Visible = bVisible
122             chkUseIndividualFont.Visible = bVisible
124             cmdSelectFont.Visible = bVisible
  
126         Case 3

128             For i = chkSMSTypeEnabled.LBound To chkSMSTypeEnabled.UBound
130                 chkSMSTypeEnabled(i).Visible = bVisible
                Next
  
132         Case 4
134             fraJobRemarks.Visible = bVisible

136             For i = txtJobRemarks.LBound To txtJobRemarks.UBound
138                 lblJobRemarks(i).Visible = bVisible
140                 txtJobRemarks(i).Visible = bVisible
                Next
  
142             fraPhonebook.Visible = bVisible

144             For i = txtPhonebookVariableField.LBound To txtPhonebookVariableField.UBound
146                 lblPhonebookVariableField(i).Visible = bVisible
148                 txtPhonebookVariableField(i).Visible = bVisible
                Next
  
150         Case 5
152             VersionSpecificAction 58, , , , bVisible
    
154             fraSingleSMS.Visible = bVisible
156             chkSingleSMSUseUserDefinedLifeTime.Visible = bVisible
158             optSingleSMSUseSpecificSettingsAsLifeTime.Visible = bVisible
160             lblSingleSMSSpecificLifeTime.Visible = bVisible
162             cboSingleSMSSpecificSettingLifeTimeUnit.Visible = bVisible
164             optSingleSMSUseSingleshotAsLifeTime.Visible = bVisible

166             fraPeriodicSMS.Visible = bVisible
168             chkPeriodicSMSUseUserDefinedLifeTime.Visible = bVisible
170             optPeriodicSMSUseWaitingTimeAsLifeTime.Visible = bVisible
172             optPeriodicSMSUseSpecificSettingsAsLifeTime.Visible = bVisible
174             lblPeriodicSMSSpecificLifeTime.Visible = bVisible
176             cboPeriodicSMSSpecificSettingLifeTimeUnit.Visible = bVisible
178             optPeriodicSMSUseSingleshotAsLifeTime.Visible = bVisible
 
        End Select

        Exit Sub

        Dim nLanguage As Integer

        '<EhFooter>
        Exit Sub

SetTabVisibility_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.SetTabVisibility " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkPeriodicSMSUseUserDefinedLifeTime_Click()
        '<EhHeader>
        On Error GoTo chkPeriodicSMSUseUserDefinedLifeTime_Click_Err
        '</EhHeader>

100     If chkPeriodicSMSUseUserDefinedLifeTime.Value = 1 Then
102         optPeriodicSMSUseWaitingTimeAsLifeTime.Enabled = True
104         optPeriodicSMSUseSpecificSettingsAsLifeTime.Enabled = True
106         optPeriodicSMSUseSingleshotAsLifeTime.Enabled = True
108         chkPeriodicSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes.Enabled = True
  
110         If optPeriodicSMSUseSpecificSettingsAsLifeTime.Value = True Then
112             txtPeriodicSMSSpecificSettingLifeTime.Enabled = True
114             cboPeriodicSMSSpecificSettingLifeTimeUnit.Enabled = True
116             lblPeriodicSMSSpecificLifeTime.Enabled = True
            Else
118             txtPeriodicSMSSpecificSettingLifeTime.Enabled = False
120             cboPeriodicSMSSpecificSettingLifeTimeUnit.Enabled = False
122             lblPeriodicSMSSpecificLifeTime.Enabled = False
            End If

        Else
124         optPeriodicSMSUseWaitingTimeAsLifeTime.Enabled = False
126         optPeriodicSMSUseSpecificSettingsAsLifeTime.Enabled = False
128         optPeriodicSMSUseSingleshotAsLifeTime.Enabled = False
130         chkPeriodicSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes.Enabled = False
132         txtPeriodicSMSSpecificSettingLifeTime.Enabled = False
134         cboPeriodicSMSSpecificSettingLifeTimeUnit.Enabled = False
136         lblPeriodicSMSSpecificLifeTime.Enabled = False
        End If

        '<EhFooter>
        Exit Sub

chkPeriodicSMSUseUserDefinedLifeTime_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.chkPeriodicSMSUseUserDefinedLifeTime_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkSingleSMSUseUserDefinedLifeTime_Click()
        '<EhHeader>
        On Error GoTo chkSingleSMSUseUserDefinedLifeTime_Click_Err
        '</EhHeader>

100     If chkSingleSMSUseUserDefinedLifeTime.Value = 1 Then
102         optSingleSMSUseSingleshotAsLifeTime.Enabled = True
104         optSingleSMSUseSpecificSettingsAsLifeTime.Enabled = True
106         chkSingleSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes.Enabled = True
  
108         If optSingleSMSUseSingleshotAsLifeTime.Value = True Then
110             txtSingleSMSSpecificSettingLifeTime.Enabled = False
112             cboSingleSMSSpecificSettingLifeTimeUnit.Enabled = False
114             lblSingleSMSSpecificLifeTime.Enabled = False
            Else
116             txtSingleSMSSpecificSettingLifeTime.Enabled = True
118             cboSingleSMSSpecificSettingLifeTimeUnit.Enabled = True
120             lblSingleSMSSpecificLifeTime.Enabled = True
            End If

        Else
122         optSingleSMSUseSpecificSettingsAsLifeTime.Enabled = False
124         optSingleSMSUseSingleshotAsLifeTime.Enabled = False
126         chkSingleSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes.Enabled = False
128         txtSingleSMSSpecificSettingLifeTime.Enabled = False
130         cboSingleSMSSpecificSettingLifeTimeUnit.Enabled = False
132         lblSingleSMSSpecificLifeTime.Enabled = False
        End If

        '<EhFooter>
        Exit Sub

chkSingleSMSUseUserDefinedLifeTime_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.chkSingleSMSUseUserDefinedLifeTime_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
        Dim frmTemp As Form
        Dim i As Integer
        Dim nCurrentIndex As Integer

100     For Each frmTemp In Forms
102         frmTemp.AdjustLanguageSettings gnLanguage
        Next

104     Unload Me
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOk_Click()
        '<EhHeader>
        On Error GoTo cmdOk_Click_Err
        '</EhHeader>
        Dim frmTemp As Form

100     If CheckScreenInput() = False Then
            Exit Sub
        End If

102     SaveGeneralSettings
104     RestoreGeneralSettings
        
        On Error Resume Next
        
106     For Each frmTemp In Forms
108         AdjustFontControls frmTemp
110         frmTemp.AdjustLanguageSettings gnLanguage
        Next

112     Set frmTemp = Nothing

114     Unload Me
        '<EhFooter>
        Exit Sub

cmdOk_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.cmdOk_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOriginatorCheck_Click()
        '<EhHeader>
        On Error GoTo cmdOriginatorCheck_Click_Err
'        '</EhHeader>
'100     frmOriginatorCheck.txtStep1Originator = txtOriginator.Text
'102     frmOriginatorCheck.Show vbModal, Me
        '<EhFooter>
        Exit Sub

cmdOriginatorCheck_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.cmdOriginatorCheck_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSelectFont_Click()
        '<EhHeader>
        On Error GoTo cmdSelectFont_Click_Err
        '</EhHeader>
        On Error GoTo CancelError
100     cmDlgFonts.Flags = cdlCFBoth
102     cmDlgFonts.ShowFont

104     lblFontPreview.caption = cmDlgFonts.FontName
106     lblFontPreview.Font.Name = cmDlgFonts.FontName
108     lblFontPreview.Font.Bold = cmDlgFonts.FontBold
110     lblFontPreview.Font.Italic = cmDlgFonts.FontItalic
112     lblFontPreview.Font.Size = cmDlgFonts.FontSize
114     lblFontPreview.Font.Strikethrough = cmDlgFonts.FontStrikethru
116     lblFontPreview.Font.Underline = cmDlgFonts.FontUnderline

        Exit Sub
CancelError:
        Exit Sub
        '<EhFooter>
        Exit Sub

cmdSelectFont_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.cmdSelectFont_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Activate()
        '<EhHeader>
        On Error GoTo Form_Activate_Err
        '</EhHeader>
100     CenterForm Me

102     Form_Load
104     tabSettings.Tab = 0
106     tabSettings_Click 0
        '<EhFooter>
        Exit Sub

Form_Activate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.Form_Activate " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        Dim i As Integer
100     mnLanguage = gnLanguage
102     VersionSpecificAction 6
104     VersionSpecificAction 8
106     AdjustFontControls Me
108     AdjustLanguageSettings gnLanguage

110     If gtFontSettingsCurrent.bSpecificFontUsed Then
112         chkUseIndividualFont.Value = 1
        Else
114         chkUseIndividualFont.Value = 0
        End If

116     txtUserkey.Text = frmSMSMain.txtUserkey.Text
118     txtPassword.Text = frmSMSMain.txtPassword.Text
120     txtOriginator.Text = frmSMSMain.txtOriginator.Text

122     If gbSaveMessagesInSendLog = True Then
124         chkSaveMessagesInSendJournal.Value = 1
        Else
126         chkSaveMessagesInSendJournal.Value = 0
        End If

128     Select Case gnLanguage

            Case 1
130             optLanguage(1).Value = True
  
132         Case 2
134             optLanguage(2).Value = True
  
136         Case Else
                'Do nothing
        End Select

138     cmDlgFonts.FontName = gtFontSettingsCurrent.sFontName
140     cmDlgFonts.FontSize = gtFontSettingsCurrent.rFontSize
142     cmDlgFonts.FontBold = gtFontSettingsCurrent.bFontBold
144     cmDlgFonts.FontItalic = gtFontSettingsCurrent.bFontItalic

146     lblFontPreview.caption = gtFontSettingsCurrent.sFontName
148     lblFontPreview.Font.Name = gtFontSettingsCurrent.sFontName
150     lblFontPreview.Font.Size = gtFontSettingsCurrent.rFontSize
152     lblFontPreview.Font.Bold = gtFontSettingsCurrent.bFontBold
154     lblFontPreview.Font.Italic = gtFontSettingsCurrent.bFontItalic

156     If gtMenuControl.bTextSMSEnabled = True Then chkSMSTypeEnabled(0).Value = 1 Else chkSMSTypeEnabled(0).Value = 0
158     If gtMenuControl.bOperatorLogoEnabled = True Then chkSMSTypeEnabled(1).Value = 1 Else chkSMSTypeEnabled(1).Value = 0
160     If gtMenuControl.bGroupLogoEnabled = True Then chkSMSTypeEnabled(2).Value = 1 Else chkSMSTypeEnabled(2).Value = 0
162     If gtMenuControl.bRingtoneEnabled = True Then chkSMSTypeEnabled(3).Value = 1 Else chkSMSTypeEnabled(3).Value = 0
164     If gtMenuControl.bPictureMessageEnabled = True Then chkSMSTypeEnabled(4).Value = 1 Else chkSMSTypeEnabled(4).Value = 0
166     If gtMenuControl.bVCardEnabled = True Then chkSMSTypeEnabled(5).Value = 1 Else chkSMSTypeEnabled(5).Value = 0
168     If gtMenuControl.bUnicodeEnabled = True Then chkSMSTypeEnabled(6).Value = 1 Else chkSMSTypeEnabled(6).Value = 0
170     If gtMenuControl.bWAPPushSMSEnabled = True Then chkSMSTypeEnabled(7).Value = 1 Else chkSMSTypeEnabled(7).Value = 0
172     If gtMenuControl.bBinaryDataEnabled = True Then chkSMSTypeEnabled(8).Value = 1 Else chkSMSTypeEnabled(8).Value = 0
174     VersionSpecificAction 9

176     For i = 1 To gcnNumberOfJobRemarkFields
178         txtJobRemarks(i).Text = GetJobRemarksFieldFromDatabase(gnLanguage, i)
        Next

180     For i = 1 To 3
182         txtPhonebookVariableField(i).Text = GetPhonebookVariableFieldFromDatabase(gnLanguage, i)
        Next

184     If gtValidityPeriodSettings.bSingleSMSUseUserDefinedLifeTime = True Then
186         chkSingleSMSUseUserDefinedLifeTime.Value = 1
        Else
188         chkSingleSMSUseUserDefinedLifeTime.Value = 0
        End If

190     If gtValidityPeriodSettings.bSingleSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes = True Then
192         chkSingleSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes.Value = 1
        Else
194         chkSingleSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes.Value = 0
        End If

196     Select Case gtValidityPeriodSettings.nSingleSMSValidityPeriodMode

            Case ValidityPeriodMode.UseSpecificSettingsAsLifeTime
198             optSingleSMSUseSpecificSettingsAsLifeTime.Value = True
200             optSingleSMSUseSingleshotAsLifeTime.Value = False
  
202         Case ValidityPeriodMode.UseSingleshotAsLifeTime
204             optSingleSMSUseSpecificSettingsAsLifeTime.Value = False
206             optSingleSMSUseSingleshotAsLifeTime.Value = True
  
        End Select
  
208     txtSingleSMSSpecificSettingLifeTime = Trim(Str$(gtValidityPeriodSettings.lSingleSMSSpecificSettingLifeTime))

210     Select Case gtValidityPeriodSettings.nSingleSMSSpecificSettingLifeTimeUnit

            Case 0
212             cboSingleSMSSpecificSettingLifeTimeUnit.ListIndex = 0
  
214         Case 1
216             cboSingleSMSSpecificSettingLifeTimeUnit.ListIndex = 1
  
        End Select

218     If gtValidityPeriodSettings.bPeriodicSMSUseUserDefinedLifeTime = True Then
220         chkPeriodicSMSUseUserDefinedLifeTime.Value = 1
        Else
222         chkPeriodicSMSUseUserDefinedLifeTime.Value = 0
        End If

224     If gtValidityPeriodSettings.bPeriodicSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes = True Then
226         chkPeriodicSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes.Value = 1
        Else
228         chkPeriodicSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes.Value = 0
        End If

230     Select Case gtValidityPeriodSettings.nPeriodicSMSValidityPeriodMode

            Case ValidityPeriodMode.UseSpecificSettingsAsLifeTime
232             optPeriodicSMSUseSpecificSettingsAsLifeTime.Value = True
234             optPeriodicSMSUseSingleshotAsLifeTime.Value = False
236             optPeriodicSMSUseWaitingTimeAsLifeTime = False
  
238         Case ValidityPeriodMode.UseSingleshotAsLifeTime
240             optPeriodicSMSUseSpecificSettingsAsLifeTime.Value = False
242             optPeriodicSMSUseSingleshotAsLifeTime.Value = True
244             optPeriodicSMSUseWaitingTimeAsLifeTime = False
  
246         Case ValidityPeriodMode.UseWaitingTimeAsLifeTime
248             optPeriodicSMSUseSpecificSettingsAsLifeTime.Value = False
250             optPeriodicSMSUseSingleshotAsLifeTime.Value = False
252             optPeriodicSMSUseWaitingTimeAsLifeTime = True
  
        End Select
  
254     txtPeriodicSMSSpecificSettingLifeTime = Trim(Str$(gtValidityPeriodSettings.lPeriodicSMSSpecificSettingLifeTime))

256     Select Case gtValidityPeriodSettings.nPeriodicSMSSpecificSettingLifeTimeUnit

            Case 0
258             cboPeriodicSMSSpecificSettingLifeTimeUnit.ListIndex = 0
  
260         Case 1
262             cboPeriodicSMSSpecificSettingLifeTimeUnit.ListIndex = 1
  
        End Select

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub optLanguage_Click(Index As Integer)
        '<EhHeader>
        On Error Resume Next
        '</EhHeader>
        Dim frmTemp As Form

100     For Each frmTemp In Forms
102         frmTemp.AdjustLanguageSettings Index
        Next

        '<EhFooter>
        Exit Sub

optLanguage_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.optLanguage_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub optPeriodicSMSUseSingleshotAsLifeTime_Click()
        '<EhHeader>
        On Error GoTo optPeriodicSMSUseSingleshotAsLifeTime_Click_Err
        '</EhHeader>
100     chkPeriodicSMSUseUserDefinedLifeTime_Click
        '<EhFooter>
        Exit Sub

optPeriodicSMSUseSingleshotAsLifeTime_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.optPeriodicSMSUseSingleshotAsLifeTime_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub optPeriodicSMSUseSpecificSettingsAsLifeTime_Click()
        '<EhHeader>
        On Error GoTo optPeriodicSMSUseSpecificSettingsAsLifeTime_Click_Err
        '</EhHeader>
100     chkPeriodicSMSUseUserDefinedLifeTime_Click
        '<EhFooter>
        Exit Sub

optPeriodicSMSUseSpecificSettingsAsLifeTime_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.optPeriodicSMSUseSpecificSettingsAsLifeTime_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub optPeriodicSMSUseWaitingTimeAsLifeTime_Click()
        '<EhHeader>
        On Error GoTo optPeriodicSMSUseWaitingTimeAsLifeTime_Click_Err
        '</EhHeader>
100     chkPeriodicSMSUseUserDefinedLifeTime_Click
        '<EhFooter>
        Exit Sub

optPeriodicSMSUseWaitingTimeAsLifeTime_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.optPeriodicSMSUseWaitingTimeAsLifeTime_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub optSingleSMSUseSingleshotAsLifeTime_Click()
        '<EhHeader>
        On Error GoTo optSingleSMSUseSingleshotAsLifeTime_Click_Err
        '</EhHeader>
100     chkSingleSMSUseUserDefinedLifeTime_Click
        '<EhFooter>
        Exit Sub

optSingleSMSUseSingleshotAsLifeTime_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.optSingleSMSUseSingleshotAsLifeTime_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub optSingleSMSUseSpecificSettingsAsLifeTime_Click()
        '<EhHeader>
        On Error GoTo optSingleSMSUseSpecificSettingsAsLifeTime_Click_Err
        '</EhHeader>
100     chkSingleSMSUseUserDefinedLifeTime_Click
        '<EhFooter>
        Exit Sub

optSingleSMSUseSpecificSettingsAsLifeTime_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.optSingleSMSUseSpecificSettingsAsLifeTime_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub tabSettings_Click(PreviousTab As Integer)
        '<EhHeader>
        On Error GoTo tabSettings_Click_Err
        '</EhHeader>
        Dim i As Integer

100     If tabSettings.Tab = 4 Then
102        SetTabVisibility tabSettings.Tab, False
        Else
104         SetTabVisibility tabSettings.Tab, True
        End If
    
106     For i = 0 To 6

108         If i <> tabSettings.Tab Then
110             SetTabVisibility i, False
            End If

        Next

        '<EhFooter>
        Exit Sub

tabSettings_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmSettings.tabSettings_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtPassword_Change()
    frmSMSMain.txtPassword.Text = txtPassword.Text
End Sub

Private Sub txtUserkey_Change()
    frmSMSMain.txtUserkey.Text = txtUserkey.Text
End Sub
