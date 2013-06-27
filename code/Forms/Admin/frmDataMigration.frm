VERSION 5.00
Begin VB.Form frmDataMigration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS  Incident Data Migration"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7755
   Icon            =   "frmDataMigration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStartMigration 
      Caption         =   "Start Migration"
      Height          =   405
      Left            =   6480
      TabIndex        =   2
      Top             =   660
      Width           =   1245
   End
   Begin VB.TextBox txtMDB 
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   270
      Width           =   7185
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   7290
      TabIndex        =   0
      Top             =   270
      Width           =   435
   End
   Begin VB.Label lblOASISClient 
      Caption         =   "OASIS Client Database To Migrate Data:"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   30
      Width           =   3255
   End
End
Attribute VB_Name = "frmDataMigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
Dim c As New cCommonDialog

    With c
        .Filter = "*.mdb"
        .DefaultExt = ".mdb"
        .DialogTitle = "Choose OASIS Client Db to Migrate"
        .InitDir = App.Path
        .ShowOpen
        txtMDB.Text = .Filename
    End With

End Sub

Private Sub cmdStartMigration_Click()
        '<EhHeader>
        On Error GoTo cmdStartMigration_Click_Err
        '</EhHeader>
        Dim cn As New ADODB.Connection
        Dim oRS As New ADODB.Recordset
        Dim oRSSynch As New ADODB.Recordset
        Dim bExists As Boolean
        Dim i As Integer
    
100     cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtMDB.Text & ";Persist Security Info=False"
    
102     oRSSynch.Open "SELECT [sID] FROM SynchHistory", cn, adOpenDynamic, adLockReadOnly
    
104     With oRS
106         .CursorLocation = adUseClient
108         .CursorType = adOpenDynamic
110         .LockType = adLockBatchOptimistic
    
112         .Open "SELECT [ID] FROM oincidents_FEA", cn
    
114         If Not .EOF And Not .BOF Then
116             .MoveFirst
            
118             Do While Not .EOF

120                 If Not oRSSynch.BOF And Not oRSSynch.EOF Then
122                     oRSSynch.MoveFirst
124                     oRSSynch.Filter = "sID = '" & .Fields.Item("ID").Value & "'"
                    
126                     If Not oRSSynch.BOF And Not oRSSynch.EOF Then
128                         bExists = True
                        End If
                    
130                     oRSSynch.Filter = adFilterNone
                    
                    End If
                
132                 If Not bExists Then
134                     SetNewSynchDBElement cn, GetGuid, .Fields.Item("ID").Value, "Synched Incident", "", "OASIS Ver 1 Data Migration ", RFC3339DateTime, "oincidents", True, "'true'"
136                     i = i + 1
                    End If
                
138                 bExists = False
140                 .MoveNext
                Loop
            
            End If
                    
142         If i > 0 Then
144             MsgBox i & " Records were migrated out of " & .RecordCount, vbInformation, "Incident migration Completed"
            Else
146             MsgBox "It seems like all data is already migrated properly. No records were migrated", vbInformation, "Incident migration Completed"
            End If
                    
        End With
    
        '<EhFooter>
        Exit Sub

cmdStartMigration_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISDataMigration.frmDataMigration.cmdStartMigration_Click " & _
               "at line " & Erl
        Exit Sub
        '</EhFooter>
End Sub
