VERSION 5.00
Begin VB.Form frmFilterJoblog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter"
   ClientHeight    =   3660
   ClientLeft      =   4725
   ClientTop       =   3600
   ClientWidth     =   5655
   Icon            =   "frmFilterJoblog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboSendingType 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2640
      Width           =   3135
   End
   Begin VB.ComboBox cboSMSType 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2280
      Width           =   3135
   End
   Begin VB.ComboBox cboJobRemarks 
      Height          =   315
      Index           =   6
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1920
      Width           =   3135
   End
   Begin VB.ComboBox cboJobRemarks 
      Height          =   315
      Index           =   5
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1560
      Width           =   3135
   End
   Begin VB.ComboBox cboJobRemarks 
      Height          =   315
      Index           =   4
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   3135
   End
   Begin VB.ComboBox cboJobRemarks 
      Height          =   315
      Index           =   3
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   3135
   End
   Begin VB.ComboBox cboJobRemarks 
      Height          =   315
      Index           =   2
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.ComboBox cboJobRemarks 
      Height          =   315
      Index           =   1
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblSMSType 
      Caption         =   "lblSMSType"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblSendingType 
      Caption         =   "lblSendingType"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lblJobRemarks 
      Caption         =   "lblJobRemarks"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblJobRemarks 
      Caption         =   "lblJobRemarks"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblJobRemarks 
      Caption         =   "lblJobRemarks"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblJobRemarks 
      Caption         =   "lblJobRemarks"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblJobRemarks 
      Caption         =   "lblJobRemarks"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblJobRemarks 
      Caption         =   "lblJobRemarks"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmFilterJoblog"
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
               "in OASIS_SMS_Messenger.frmFilterJoblog.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOk_Click()
'        '<EhHeader>
'        On Error GoTo cmdOk_Click_Err
'        '</EhHeader>
'    Dim i As Integer
'
'100 For i = 1 To gcnNumberOfJobRemarkFields
'102   If cboJobRemarks(i).ListIndex <> -1 Then
'104     If cboJobRemarks(i).ItemData(cboJobRemarks(i).ListIndex) <> gcNoFilter Then
'106       gtJobLogFilterSettings.bJobRemarksFilterUsed(i) = True
'108       gtJobLogFilterSettings.sJobRemarks(i) = cboJobRemarks(i).List(cboJobRemarks(i).ListIndex)
'        Else
'110       gtJobLogFilterSettings.bJobRemarksFilterUsed(i) = False
'112       gtJobLogFilterSettings.sJobRemarks(i) = ""
'        End If
'      End If
'    Next
'
'114 If cboSMSType.ListIndex <> -1 Then
'116   If cboSMSType.ItemData(cboSMSType.ListIndex) <> gcNoFilter Then
'118     gtJobLogFilterSettings.bSMSTypeFilterUsed = True
'120     gtJobLogFilterSettings.sSMSType = cboSMSType.List(cboSMSType.ListIndex)
'      Else
'122     gtJobLogFilterSettings.bSMSTypeFilterUsed = False
'124     gtJobLogFilterSettings.sSMSType = ""
'      End If
'    End If
'
'126 If cboSendingType.ListIndex <> -1 Then
'128   If cboSendingType.ItemData(cboSendingType.ListIndex) <> gcNoFilter Then
'130     gtJobLogFilterSettings.bSendingTypeFilterUsed = True
'132     gtJobLogFilterSettings.sSendingType = cboSendingType.List(cboSendingType.ListIndex)
'      Else
'134     gtJobLogFilterSettings.bSendingTypeFilterUsed = False
'136     gtJobLogFilterSettings.sSendingType = ""
'      End If
'    End If
'
'138 Call frmJoblog.FormLoadWithoutSubClassing
'140 Unload Me
'        '<EhFooter>
'        Exit Sub
'
'cmdOk_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmFilterJoblog.cmdOk_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub
Private Sub Form_Load()
'        '<EhHeader>
'        On Error GoTo Form_Load_Err
'        '</EhHeader>
'    Dim sSQL As String
'    Dim i As Integer
'100 CenterForm Me
'    On Error GoTo ErrorTrap:
'102 AdjustFontControls Me
'
'104 AdjustLanguageSettings gnLanguage
'
'106 For i = 1 To gcnNumberOfJobRemarkFields
'108   If gtJobLogFilterSettings.bJobRemarksFilterUsed(i) Then
'110     SetListIndexWithText cboJobRemarks(i), gtJobLogFilterSettings.sJobRemarks(i)
'      Else
'112     SetListIndexWithItemData cboJobRemarks(i), gcNoFilter
'      End If
'    Next
'
'114 If gtJobLogFilterSettings.bSMSTypeFilterUsed Then
'116   SetListIndexWithText cboSMSType, gtJobLogFilterSettings.sSMSType
'    Else
'118   SetListIndexWithItemData cboSMSType, gcNoFilter
'    End If
'
'120 If gtJobLogFilterSettings.bSendingTypeFilterUsed Then
'122   SetListIndexWithText cboSendingType, gtJobLogFilterSettings.sSendingType
'    Else
'124   SetListIndexWithItemData cboSendingType, gcNoFilter
'    End If
'
'    Exit Sub
'ErrorTrap:
'    Exit Sub
'
'        '<EhFooter>
'        Exit Sub
'
'Form_Load_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmFilterJoblog.Form_Load " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub
Public Sub AdjustLanguageSettings(nLanguage As Integer)
'        '<EhHeader>
'        On Error GoTo AdjustLanguageSettings_Err
'        '</EhHeader>
'    Dim i As Integer
'100 cmdOk.caption = LoadLanguageSpecificString(nLanguage, 1)
'102 cmdCancel.caption = LoadLanguageSpecificString(nLanguage, 2)
'
'    'Jobremarks
'104 For i = 1 To gcnNumberOfJobRemarkFields
'106   UpdateComboBoxfrmFilterJoblog i, Me.cboJobRemarks(i), "JobRemarksField"
'    Next
'
'108 For i = 1 To gcnNumberOfJobRemarkFields
'110   lblJobRemarks(i).caption = GetJobRemarksFieldFromDatabase(nLanguage, i)
'    Next
'
'112 lblSMSType.caption = LoadLanguageSpecificString(nLanguage, 570)
'114 lblSendingType.caption = LoadLanguageSpecificString(nLanguage, 571)
'
'116 UpdateComboBoxfrmFilterJoblog -1, cboSMSType, "FSMSType"
'118 UpdateComboBoxfrmFilterJoblog -1, cboSendingType, "FSendingType"
'        '<EhFooter>
'        Exit Sub
'
'AdjustLanguageSettings_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmFilterJoblog.AdjustLanguageSettings " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub
