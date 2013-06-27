VERSION 5.00
Begin VB.Form frmAccessTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OASIS Database utilities"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmAccessTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Form"
      Height          =   345
      Left            =   180
      TabIndex        =   19
      Top             =   5040
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   5715
      TabIndex        =   18
      Top             =   5040
      Width           =   1200
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "Show About"
      Height          =   345
      Left            =   3480
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Frame fraTestAccessDB 
      Caption         =   "OASIS Client Database Utilities "
      Height          =   4755
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtDBRestore 
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   3720
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetRestorePath 
         Height          =   375
         Left            =   5820
         Picture         =   "frmAccessTest.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3690
         Width           =   435
      End
      Begin VB.CommandButton cmdLookup2 
         Height          =   375
         Left            =   5820
         Picture         =   "frmAccessTest.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2950
         Width           =   435
      End
      Begin VB.TextBox txtDBBackup 
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   2980
         Width           =   5415
      End
      Begin VB.CommandButton cmdRestoreAccess 
         Caption         =   "Restore Database"
         Height          =   375
         Left            =   3780
         TabIndex        =   9
         Top             =   4140
         Width           =   1935
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Compact/Repair"
         Height          =   375
         Left            =   3780
         TabIndex        =   8
         Top             =   1080
         Width           =   1875
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy DB (Backup)"
         Height          =   375
         Left            =   3780
         TabIndex        =   7
         Top             =   1980
         Width           =   1875
      End
      Begin VB.TextBox txtCopyDB 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   1560
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetFolder 
         Height          =   375
         Left            =   5820
         Picture         =   "frmAccessTest.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1530
         Width           =   435
      End
      Begin VB.TextBox txtDB 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   5415
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   375
         Left            =   5820
         Picture         =   "frmAccessTest.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   630
         Width           =   435
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   435
         Left            =   180
         TabIndex        =   14
         Top             =   4140
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.Label lblTest 
         Caption         =   "Enter Path to Restore DB To (only path - no file name)"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   3380
         Width           =   5415
      End
      Begin VB.Line linSep 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   180
         X2              =   6540
         Y1              =   2480
         Y2              =   2480
      End
      Begin VB.Line linSep 
         Index           =   0
         X1              =   180
         X2              =   6540
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Label lblTest 
         Caption         =   "Enter Path of Backup DB to Restore (path and file name)"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   2640
         Width           =   5415
      End
      Begin VB.Label llblTest 
         Caption         =   "Enter Path to Backup DB To"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   1260
         Width           =   5415
      End
      Begin VB.Label lblTest 
         Caption         =   "Enter Path to Access DB"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmAccessTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_objMaintain As cMaintainDB
Attribute m_objMaintain.VB_VarHelpID = -1

Private Const m_MESSAGE As String = "Please enter required data in fields."

Private Sub cmdAbout_Click()
        '<EhHeader>
        On Error GoTo cmdAbout_Click_Err
        '</EhHeader>

100     m_objMaintain.About

        '<EhFooter>
        Exit Sub

cmdAbout_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.cmdAbout_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdClear_Click()
        '<EhHeader>
        On Error GoTo cmdClear_Click_Err
        '</EhHeader>
    
100     txtCopyDB.Text = ""
102     txtDB.Text = ""
104     txtDBBackup.Text = ""
106     txtDBRestore.Text = ""
108     lblmsg.Visible = False
    
        '<EhFooter>
        Exit Sub

cmdClear_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.cmdClear_Click " & _
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
               "in OASISClient.frmAccessTest.cmdClose_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCopy_Click()
        '<EhHeader>
        On Error GoTo cmdCopy_Click_Err
        '</EhHeader>

        Dim strOriginalDB As String
        Dim strCopyPath As String
        Dim bSuccess As Boolean
        Dim objMouse As cMouseCursor

100     Set objMouse = New cMouseCursor
102     objMouse.SetCursor vbHourglass

104     If txtDB.Text = "" Then
106         lblmsg.Visible = True
108         txtDB.SetFocus
            Exit Sub
        Else
110         lblmsg.Visible = False
        End If
    
112     If txtCopyDB.Text = "" Then
114         txtCopyDB.SetFocus
116         lblmsg.Visible = True
            Exit Sub
        Else
118         lblmsg.Visible = False
        End If

120     strOriginalDB = txtDB.Text
122     strCopyPath = txtCopyDB.Text
    
        ' now we have the info needed to do the backup
124     bSuccess = m_objMaintain.BackupAccessDB(strOriginalDB, strCopyPath)
126     If bSuccess Then
128         MsgBox "DB copied to: " & strCopyPath
        End If

130     DebugPrint "DB Name: " & m_objMaintain.DBName
132     DebugPrint "DB Path: " & m_objMaintain.DBPath

        '<EhFooter>
        Exit Sub

cmdCopy_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.cmdCopy_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdGetRestorePath_Click()
        '<EhHeader>
        On Error GoTo cmdGetRestorePath_Click_Err
        '</EhHeader>
        Dim strFile As String
        Dim strPrompt As String

100     strPrompt = "Select the folder to copy your backup file to."

102     strFile = BrowseForFolder(strPrompt, Me.hWnd)

104     txtDBRestore.Text = strFile
        '<EhFooter>
        Exit Sub

cmdGetRestorePath_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.cmdGetRestorePath_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdLookup_Click()
        '<EhHeader>
        On Error GoTo cmdLookup_Click_Err
        '</EhHeader>

        On Error Resume Next
    
        Dim strFile As String
        Dim c As New cCommonDialog
    
100     With c
102         .CancelError = True
104         .DialogTitle = "Select Access DB to Compress"
106         .Filter = "Access Database(*.mdb) | *.mdb"
108         .InitDir = g_sAppPath
110         .ShowOpen
        End With

112     If Err.number <> 0 Then
            ' user cancelled
114         strFile = ""
116         Me.txtDB.Text = ""
118         txtDB.SetFocus
            Exit Sub
        End If

120     strFile = c.Filename

122     txtDB.Text = strFile

        '<EhFooter>
        Exit Sub

cmdLookup_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.cmdLookup_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdLookup2_Click()
        '<EhHeader>
        On Error GoTo cmdLookup2_Click_Err
        '</EhHeader>
    
        On Error Resume Next
    
        Dim strFile As String
        Dim c As New cCommonDialog
    
100     With c
102         .CancelError = True
104         .DialogTitle = "Select Database to restore"
106         .Filter = "Access Database(*.mdb) | *.mdb"
108         .InitDir = g_sAppPath
110         .ShowOpen
        
        End With

112     If Err.number <> 0 Then
            ' user cancelled
114         strFile = ""
116         txtDBBackup.Text = ""
118         txtDBBackup.SetFocus
            Exit Sub
        End If

120     strFile = c.Filename

122     txtDBBackup.Text = strFile

        '<EhFooter>
        Exit Sub

cmdLookup2_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.cmdLookup2_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdRestoreAccess_Click()
        '<EhHeader>
        On Error GoTo cmdRestoreAccess_Click_Err
        '</EhHeader>

        Dim strBackupPath As String
        Dim strRestorePath As String

100     strBackupPath = txtDBBackup.Text
102     strRestorePath = txtDBRestore.Text

104     If strBackupPath = "" Then
106         txtDBBackup.SetFocus
108         lblmsg.Visible = True
            Exit Sub
        Else
110         lblmsg.Visible = False
        End If

112     If strRestorePath = "" Then
114         txtDBRestore.SetFocus
116         lblmsg.Visible = True
            Exit Sub
        Else
118         lblmsg.Visible = False
        End If

120     m_objMaintain.RestoreAccessDB strBackupPath, strRestorePath, True

        '<EhFooter>
        Exit Sub

cmdRestoreAccess_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.cmdRestoreAccess_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTest_Click()
        '<EhHeader>
        On Error GoTo cmdTest_Click_Err
        '</EhHeader>

        Dim strOriginalDB As String
        Dim bSuccess As Boolean
        Dim objMouse As cMouseCursor

100     Set objMouse = New cMouseCursor
102     objMouse.SetCursor vbHourglass

104     strOriginalDB = txtDB.Text

106     If strOriginalDB = "" Then
108         lblmsg.Visible = True
110         txtDB.SetFocus
            Exit Sub
        Else
112         lblmsg.Visible = False
        End If

114     bSuccess = m_objMaintain.CompactAccessDB(strOriginalDB)

        '<EhFooter>
        Exit Sub

cmdTest_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.cmdTest_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub cmdGetFolder_Click()
        '<EhHeader>
        On Error GoTo cmdGetFolder_Click_Err
        '</EhHeader>

        Dim strFile As String
        Dim strPrompt As String

100     strPrompt = "Select the folder to copy your backup file to."

102     strFile = BrowseForFolder(strPrompt, Me.hWnd)

104     txtCopyDB.Text = strFile

        '<EhFooter>
        Exit Sub

cmdGetFolder_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.cmdGetFolder_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

        On Error Resume Next

100     Set m_objMaintain = New cMaintainDB
102     lblmsg.caption = m_MESSAGE
    
104     If Not g_sLanguage = "" Then
106         If Not m_Cnn.State = adStateClosed Then
108             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If

        'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Violations;PWD="";Data Source=(Local)

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

        On Error Resume Next

100     Set m_objMaintain = Nothing

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_objMaintain_BackupError(ByVal FailMessage As String, ByVal BackupFolder As String)
        '<EhHeader>
        On Error GoTo m_objMaintain_BackupError_Err
        '</EhHeader>

        Dim strMsg As String

100     strMsg = FailMessage
102     strMsg = strMsg & vbCrLf & "Backup folder = " & BackupFolder

104     MsgBox strMsg, vbOKOnly + vbInformation, "Backup Failed"


        '<EhFooter>
        Exit Sub

m_objMaintain_BackupError_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.m_objMaintain_BackupError " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_objMaintain_BackupFinished(ByVal BackupFolder As String)
        '<EhHeader>
        On Error GoTo m_objMaintain_BackupFinished_Err
        '</EhHeader>

        Dim strMsg As String

100     strMsg = "The backup completed successfully. The files are located at: " & BackupFolder

102     MsgBox strMsg

        '<EhFooter>
        Exit Sub

m_objMaintain_BackupFinished_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.m_objMaintain_BackupFinished " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_objMaintain_BackupLeftOnServer(ByVal FailMessage As String, ByVal FileOnServer As String)
        '<EhHeader>
        On Error GoTo m_objMaintain_BackupLeftOnServer_Err
        '</EhHeader>

100     MsgBox FailMessage & vbCrLf & "File is: " & FileOnServer, vbOKOnly + vbInformation, "Cannot copy backup file"

        '<EhFooter>
        Exit Sub

m_objMaintain_BackupLeftOnServer_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.m_objMaintain_BackupLeftOnServer " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_objMaintain_CompactError(ByVal OriginalDB As String, ByVal DBCopy As String, ByVal FailMessage As String)
        '<EhHeader>
        On Error GoTo m_objMaintain_CompactError_Err
        '</EhHeader>

        Dim strMsg As String

100     strMsg = "There was a problem with compacting the database."
102     strMsg = strMsg & vbCrLf & "For your information, the original path is: "
104     strMsg = strMsg & vbCrLf & OriginalDB
106     strMsg = strMsg & vbCrLf & "The DB was first copied to the following location: "
108     strMsg = strMsg & vbCrLf & DBCopy
110     strMsg = strMsg & vbCrLf & "It is possible that the original DB is missing or corrupted. If so, "
112     strMsg = strMsg & vbCrLf & "The copy should be good. In this case, replace the original with the copy."
114     strMsg = strMsg & vbCrLf & "The error is reported as: "
116     strMsg = strMsg & vbCrLf & FailMessage

118     MsgBox strMsg, vbOKOnly + vbCritical, "Error in compacting DB"

        '<EhFooter>
        Exit Sub

m_objMaintain_CompactError_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.m_objMaintain_CompactError " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_objMaintain_CompactFinished(ByVal OriginalDB As String, ByVal DBCopy As String)
        '<EhHeader>
        On Error GoTo m_objMaintain_CompactFinished_Err
        '</EhHeader>

        Dim strMsg As String

100     strMsg = "The database was successfully compacted/repaired."
102     strMsg = strMsg & vbCrLf & "The compacted DB = " & OriginalDB

104     MsgBox strMsg

        '<EhFooter>
        Exit Sub

m_objMaintain_CompactFinished_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.m_objMaintain_CompactFinished " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_objMaintain_CopyCancelled()
        '<EhHeader>
        On Error GoTo m_objMaintain_CopyCancelled_Err
        '</EhHeader>

100     MsgBox "User cancelled the database copy process."

        '<EhFooter>
        Exit Sub

m_objMaintain_CopyCancelled_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.m_objMaintain_CopyCancelled " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_objMaintain_CopyError(ByVal OriginalPath As String, ByVal CopyPath As String, ByVal FailMessage As String)
        '<EhHeader>
        On Error GoTo m_objMaintain_CopyError_Err
        '</EhHeader>

        Dim strMsg As String

100     strMsg = "There was a problem with copying the database for backup or maintenance purposes."
102     strMsg = strMsg & vbCrLf & "For your information, the original path is: "
104     strMsg = strMsg & vbCrLf & OriginalPath
106     strMsg = strMsg & vbCrLf & "The path to copy the DB to is: "
108     strMsg = strMsg & vbCrLf & CopyPath
110     strMsg = strMsg & vbCrLf & "The error is reported as: "
112     strMsg = strMsg & vbCrLf & FailMessage

114     MsgBox strMsg, vbOKOnly + vbCritical, "Error in copying DB"

        '<EhFooter>
        Exit Sub

m_objMaintain_CopyError_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.m_objMaintain_CopyError " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_objMaintain_CreateAFolder(ByVal FolderPath As String, Cancel As Boolean)
        '<EhHeader>
        On Error GoTo m_objMaintain_CreateAFolder_Err
        '</EhHeader>

        Dim intResponse As Integer
        Dim strMsg As String

100     strMsg = "The folder " & FolderPath & " does not exist. Do you want to create it?"
102     strMsg = strMsg & vbCrLf & "Click on Yes to create the folder, No if you don't want to create it."
104     strMsg = strMsg & vbCrLf & "If you click on No, the database will NOT be copied."
106     intResponse = MsgBox(strMsg, vbYesNo + vbQuestion, "Create Folder?")
108     If intResponse = vbNo Then
110         Cancel = True
        End If

        '<EhFooter>
        Exit Sub

m_objMaintain_CreateAFolder_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.m_objMaintain_CreateAFolder " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_objMaintain_RestoreError(ByVal FailMessage As String)
        '<EhHeader>
        On Error GoTo m_objMaintain_RestoreError_Err
        '</EhHeader>

100     MsgBox FailMessage, vbOKOnly + vbCritical, "Unable to restore DB"

        '<EhFooter>
        Exit Sub

m_objMaintain_RestoreError_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.m_objMaintain_RestoreError " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_objMaintain_RestoreFinished(ByVal Message As String, Database As String)
        '<EhHeader>
        On Error GoTo m_objMaintain_RestoreFinished_Err
        '</EhHeader>

100     MsgBox Message, vbOKOnly + vbInformation, "Restore Successful!"

        '<EhFooter>
        Exit Sub

m_objMaintain_RestoreFinished_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAccessTest.m_objMaintain_RestoreFinished " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
