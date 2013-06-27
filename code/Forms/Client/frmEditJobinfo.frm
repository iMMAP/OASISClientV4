VERSION 5.00
Begin VB.Form frmEditJobInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Jobinfo"
   ClientHeight    =   5175
   ClientLeft      =   4320
   ClientTop       =   3255
   ClientWidth     =   5910
   Icon            =   "frmEditJobinfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraJobRemarks 
      Caption         =   "Remarks"
      Height          =   4455
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox cboJobRemarks 
         Height          =   315
         Index           =   6
         Left            =   2400
         TabIndex        =   11
         Top             =   2160
         Width           =   3135
      End
      Begin VB.ComboBox cboJobRemarks 
         Height          =   315
         Index           =   5
         Left            =   2400
         TabIndex        =   9
         Top             =   1800
         Width           =   3135
      End
      Begin VB.ComboBox cboJobRemarks 
         Height          =   315
         Index           =   4
         Left            =   2400
         TabIndex        =   7
         Top             =   1440
         Width           =   3135
      End
      Begin VB.ComboBox cboJobRemarks 
         Height          =   315
         Index           =   3
         Left            =   2400
         TabIndex        =   5
         Top             =   1080
         Width           =   3135
      End
      Begin VB.ComboBox cboJobRemarks 
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   3135
      End
      Begin VB.ComboBox cboJobRemarks 
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox txtJobRemarksMemo 
         Height          =   1815
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label lblJobRemarksMemo 
         Caption         =   "lblJobRemarksMemo"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblJobRemarks 
         Caption         =   "lblJobRemarks"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblJobRemarks 
         Caption         =   "lblJobRemarks"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblJobRemarks 
         Caption         =   "lblJobRemarks"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblJobRemarks 
         Caption         =   "lblJobRemarks"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblJobRemarks 
         Caption         =   "lblJobRemarks"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblJobRemarks 
         Caption         =   "lblJobRemarks"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   2175
   End
End
Attribute VB_Name = "frmEditJobInfo"
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
    Dim i As Integer

100 Me.caption = LoadLanguageSpecificString(nLanguage, 587)
102 cmdClose.caption = LoadLanguageSpecificString(nLanguage, 624)

104 For i = 1 To gcnNumberOfJobRemarkFields
106   lblJobRemarks(i).caption = GetJobRemarksFieldFromDatabase(nLanguage, i)
    Next

108 fraJobRemarks.caption = LoadLanguageSpecificString(nLanguage, 581)
110 lblJobRemarksMemo.caption = LoadLanguageSpecificString(nLanguage, 581)

        '<EhFooter>
        Exit Sub

AdjustLanguageSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmEditJobInfo.AdjustLanguageSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub PrepareDisplay()
        '<EhHeader>
        On Error GoTo PrepareDisplay_Err
        '</EhHeader>
    Dim rsMain As Recordset
    Dim sField As String
    Dim sEntry As String
    Dim sSelectedField As String
    Dim i As Integer
    Dim sSQL As String

    On Error GoTo ErrorTrap

100 For i = 1 To gcnNumberOfJobRemarkFields
102   sField = "JobRemarksField" & Right("0" & Trim(Str$(i)), 2)

104   Set rsMain = gdbMain.OpenRecordset("select " & sField & " from Jobs group by " & sField & " order by " & sField)

106   If rsMain.BOF And rsMain.EOF Then
        'Do nothing
      Else
108     cboJobRemarks(i).Clear
110     Do While Not rsMain.EOF
112       sEntry = rsMain(sField) & ""
114       cboJobRemarks(i).AddItem sEntry
116       rsMain.MoveNext
        Loop
      End If
118 rsMain.Close
    Next

120 sSQL = "SELECT * FROM Jobs WHERE lID = " & Str$(glJobIDEditJobRemarks)

122 Set rsMain = gdbMain.OpenRecordset(sSQL)

124 If rsMain.BOF And rsMain.EOF Then
      'Do nothing
    Else
126   For i = 1 To gcnNumberOfJobRemarkFields
128     cboJobRemarks(i).Text = rsMain("JobRemarksField" & Right("0" & Trim(Str$(i)), 2)) & ""
      Next
130   txtJobRemarksMemo.Text = rsMain("JobRemarksMemo") & ""
    End If

132 rsMain.Close

    Exit Sub
ErrorTrap:
    Exit Sub

        '<EhFooter>
        Exit Sub

PrepareDisplay_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmEditJobInfo.PrepareDisplay " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub SaveJobInfo()
'        '<EhHeader>
'        On Error GoTo SaveJobInfo_Err
'        '</EhHeader>
'    Dim rsMain As Recordset
'    Dim sField As String
'    Dim sEntry As String
'    Dim sSelectedField As String
'    Dim i As Integer
'    Dim sSQL As String
'
'    On Error GoTo ErrorTrap
'
'100 sSQL = "SELECT * FROM Jobs WHERE lID = " & Str$(glJobIDEditJobRemarks)
'
'102 Set rsMain = gdbMain.OpenRecordset(sSQL)
'
'104 If rsMain.BOF And rsMain.EOF Then
'      'Do nothing
'    Else
'106   rsMain.Edit
'108   For i = 1 To gcnNumberOfJobRemarkFields
'110     rsMain("JobRemarksField" & Right("0" & Trim(Str$(i)), 2)) = Left(cboJobRemarks(i).Text, 80)
'      Next
'112   rsMain("JobRemarksMemo") = txtJobRemarksMemo.Text
'114   rsMain.UpDate
'    End If
'
'116 rsMain.Close
'
'    Exit Sub
'ErrorTrap:
'    Exit Sub
'
'
'        '<EhFooter>
'        Exit Sub
'
'SaveJobInfo_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASIS_SMS_Messenger.frmEditJobInfo.SaveJobInfo " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub


Private Sub cmdClose_Click()
        '<EhHeader>
        On Error GoTo cmdClose_Click_Err
        '</EhHeader>
100 Unload Me
        '<EhFooter>
        Exit Sub

cmdClose_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmEditJobInfo.cmdClose_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOk_Click()
        '<EhHeader>
        On Error GoTo cmdOk_Click_Err
        '</EhHeader>
'100 SaveJobInfo
'102 frmJoblog.FormLoadWithoutSubClassing
'104 Unload Me
        '<EhFooter>
        Exit Sub

cmdOk_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmEditJobInfo.cmdOk_Click " & _
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
104 PrepareDisplay
106 AdjustLanguageSettings gnLanguage
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.frmEditJobInfo.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

