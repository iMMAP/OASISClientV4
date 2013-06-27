VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmScoring 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OASIS Incidents Scoring"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4140
   Icon            =   "frmScoring.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2475
      TabIndex        =   5
      Top             =   4815
      Width           =   780
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3330
      TabIndex        =   4
      Top             =   4815
      Width           =   780
   End
   Begin C1SizerLibCtl.C1Tab c1Tab 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4110
      _cx             =   7250
      _cy             =   8281
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "Type|Target|Time"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin DXDBGRIDLibCtl.dxDBGrid dxTypeScoring 
         Height          =   4320
         Left            =   45
         OleObjectBlob   =   "frmScoring.frx":6852
         TabIndex        =   1
         Top             =   330
         Width           =   4020
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxTargetScoring 
         Height          =   4320
         Left            =   4755
         OleObjectBlob   =   "frmScoring.frx":7FB7
         TabIndex        =   2
         Top             =   330
         Width           =   4020
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxTimeScoring 
         Height          =   4320
         Left            =   5055
         OleObjectBlob   =   "frmScoring.frx":9704
         TabIndex        =   3
         Top             =   330
         Width           =   4020
      End
   End
   Begin VB.Frame FraActivate 
      Caption         =   "Activate:"
      Height          =   510
      Left            =   0
      TabIndex        =   6
      Top             =   4680
      Width           =   2355
      Begin VB.CheckBox chkChkActivate 
         Caption         =   "Time"
         Height          =   285
         Index           =   2
         Left            =   1620
         TabIndex        =   9
         Top             =   180
         Value           =   1  'Checked
         Width           =   690
      End
      Begin VB.CheckBox chkChkActivate 
         Caption         =   "Target"
         Height          =   285
         Index           =   1
         Left            =   810
         TabIndex        =   8
         Top             =   180
         Value           =   1  'Checked
         Width           =   1140
      End
      Begin VB.CheckBox chkChkActivate 
         Caption         =   "Type"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   180
         Value           =   1  'Checked
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmScoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event SetNewScoring()
Public m_bApply As Boolean

Public Sub Init()
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
100     m_bApply = False

102     With dxTypeScoring
104         .Dataset.ADODataset.ConnectionString = m_Cnn.ConnectionString
106         .Dataset.Open
108         .Dataset.Active = True
        End With

110     With dxTargetScoring
112         .Dataset.ADODataset.ConnectionString = m_Cnn.ConnectionString
114         .Dataset.Open
116         .Dataset.Active = True
        End With

118     With dxTimeScoring
120         .Dataset.ADODataset.ConnectionString = m_Cnn.ConnectionString
122         .Dataset.Open
124         .Dataset.Active = True
        End With

    'Dim rs As New ADODB.Recordset
    '
    ''    rs.Open "SELECT INCIDENT_TYPE_ID, Scoring FROM incTypeCategory", m_Cnn
    '
    '    With dxTypeScoring.Dataset.ADODataset
    '        .ConnectionString = m_Cnn.ConnectionString
    '        .CommandType = cmdText
    '        .CommandText = "SELECT Incident_Type_Name, Scoring FROM incTypeCategory"
    '        .CursorType = ctKeyset
    '        .LockType = ltBatchOptimistic
    '        .Requery
    '    End With
    '
    '    With dxTargetScoring.ADODataset
    '        .ConnectionString = m_Cnn.ConnectionString
    '        .CommandType = cmdText
    '        .CommandText = "SELECT Name, Scoring FROM incTargetCategory"
    '        .CursorType = ctKeyset
    '        .LockType = ltBatchOptimistic
    '        .Requery
    '    End With
    '
    '    With dxTimeScoring.Dataset.ADODataset
    '        .ConnectionString = m_Cnn.ConnectionString
    '        .CommandType = cmdText
    '        .CommandText = "SELECT Incident_Time_Name, Scoring FROM incTimeCategory"
    '        .CursorType = ctKeyset
    '        .LockType = ltBatchOptimistic
    '        .Requery
    '    End With
    '
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmScoring.init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdApply_Click()
    m_bApply = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    m_bApply = False
    Me.Hide
End Sub

Private Sub Form_Load()
If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
End Sub
