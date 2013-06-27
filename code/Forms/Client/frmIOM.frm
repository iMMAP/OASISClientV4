VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmIOMJOC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IOM JOC Explorer"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12480
   Icon            =   "frmIOM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIOM.frx":6852
   ScaleHeight     =   6570
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   45
      Picture         =   "frmIOM.frx":D0A4
      ScaleHeight     =   750
      ScaleWidth      =   1605
      TabIndex        =   7
      Top             =   5760
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tools"
      Height          =   780
      Left            =   1755
      TabIndex        =   1
      Top             =   5760
      Width           =   10680
      Begin VB.ComboBox ComTheme 
         Height          =   315
         ItemData        =   "frmIOM.frx":DCD6
         Left            =   3555
         List            =   "frmIOM.frx":DCE0
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdMapCharts 
         Caption         =   "Map Charts"
         Height          =   465
         Left            =   1800
         TabIndex        =   6
         Top             =   225
         Width           =   1635
      End
      Begin VB.CommandButton cmdMapThematics 
         Caption         =   "Activate Thematics"
         Height          =   465
         Left            =   5445
         TabIndex        =   5
         Top             =   225
         Width           =   1635
      End
      Begin VB.CommandButton cmdShowIn 
         Caption         =   "Show In Map"
         Height          =   465
         Left            =   7245
         TabIndex        =   4
         Top             =   225
         Width           =   1635
      End
      Begin VB.CommandButton cmdUpdateServer 
         Caption         =   "Synchronize"
         Height          =   465
         Left            =   8955
         TabIndex        =   3
         Top             =   225
         Width           =   1635
      End
      Begin VB.CommandButton cmdExportTo 
         Caption         =   "Export to Excel"
         Height          =   465
         Left            =   90
         TabIndex        =   2
         Top             =   225
         Width           =   1635
      End
      Begin VB.Label lblChooseTheme 
         Caption         =   "Choose Theme:"
         Height          =   195
         Left            =   3555
         TabIndex        =   9
         Top             =   135
         Width           =   1725
      End
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxIOM 
      Height          =   5730
      Left            =   0
      OleObjectBlob   =   "frmIOM.frx":DD12
      TabIndex        =   0
      Top             =   45
      Width           =   12435
   End
End
Attribute VB_Name = "frmIOMJOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ShowInMap(sID As String)
Public Event createChart(sName As String)
Public Event CreateTheme(sName As String)

Public Sub Init()
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
100     With dxIOM
    
102         If .Dataset.Active Then .Dataset.Close
        
104          .Dataset.ADODataset.ConnectionString = m_Cnn.ConnectionString ' "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\IM.mdb;Persist Security Info=False"
    '         .Dataset.ADODataset.CommandText = g_RSAppSettings.Fields.Item("SettingValue1").Value '"SELECT ID, Impact, Town, Province, District FROM Scoring ORDER BY Scoring DESC"
    '         .Dataset.ADODataset.CommandType = cmdText
106          .Dataset.Open
        
        End With

        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmIOMJOC.init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub cmdExportTo_Click()
        '<EhHeader>
        On Error GoTo cmdExportTo_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog
    
100     c.ShowSave
102     dxIOM.m.ExportToXLS c.Filename
104     ShellExecute Me.hwnd, "", c.Filename, "", "", 0
        '<EhFooter>
        Exit Sub

cmdExportTo_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmIOMJOC.cmdExportTo_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdMapCharts_Click()
        '<EhHeader>
        On Error GoTo cmdMapCharts_Click_Err
        '</EhHeader>
100     MsgBox "Not Implemented Yet!"
        '<EhFooter>
        Exit Sub

cmdMapCharts_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmIOMJOC.cmdMapCharts_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdMapThematics_Click()
        '<EhHeader>
        On Error GoTo cmdMapThematics_Click_Err
        '</EhHeader>
100     Select Case ComTheme.ListIndex
    
            Case 0
102             RaiseEvent CreateTheme("IOMW3")
104         Case 1
106             RaiseEvent CreateTheme("IOMFAM")
108         Case Else
110             MsgBox "Choose Theme First...", vbInformation, "IOM JOC Explorer"
        End Select
        '<EhFooter>
        Exit Sub

cmdMapThematics_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmIOMJOC.cmdMapThematics_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdShowIn_Click()
        '<EhHeader>
        On Error GoTo cmdShowIn_Click_Err
        '</EhHeader>
    Dim s1 As String

100     With dxIOM

102         If Not .Columns.FocusedColumn Is Nothing Then
104             If Not .ex.FocusedNode Is Nothing Then
                    'Get The ID
                    's1 = .Ex.FocusedNode.Strings(0)
106                 s1 = .ex.FocusedNode.Strings(0)
                    's2 = .Columns.FocusedColumn.Value
                    's1 = .Dataset.FieldValues("ID")
                End If
            End If

        End With

108     RaiseEvent ShowInMap(s1)
        '<EhFooter>
        Exit Sub

cmdShowIn_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmIOMJOC.cmdShowIn_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdUpdateServer_Click()
        '<EhHeader>
        On Error GoTo cmdUpdateServer_Click_Err
        '</EhHeader>
100     MsgBox "Server Not Available for Synchronisation", vbInformation, "IOM JOC Explorer"
        '<EhFooter>
        Exit Sub

cmdUpdateServer_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmIOMJOC.cmdUpdateServer_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     If Not g_sLanguage = "" Then
102         If Not m_Cnn.State = adStateClosed Then
104             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmIOMJOC.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
