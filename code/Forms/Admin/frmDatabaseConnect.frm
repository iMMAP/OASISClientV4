VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmDatabaseConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Administration Toolbox"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8265
   Icon            =   "frmDatabaseConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Elastic scNavigator 
      Height          =   4620
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8265
      _cx             =   14579
      _cy             =   8149
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
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   5292196
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   0
      BorderWidth     =   3
      ChildSpacing    =   2
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.PictureBox PictureDialogLarge 
         Height          =   255
         Left            =   5280
         Picture         =   "frmDatabaseConnect.frx":6852
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox PictureDialogSmall 
         Height          =   255
         Left            =   4320
         Picture         =   "frmDatabaseConnect.frx":1830F
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame optOperation 
         BackColor       =   &H0050C0A4&
         Caption         =   "Select Operation"
         ForeColor       =   &H8000000E&
         Height          =   1935
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   8055
         Begin VB.OptionButton optArray 
            BackColor       =   &H0050C0A4&
            Caption         =   "Database Explorer"
            Height          =   495
            Index           =   3
            Left            =   3120
            TabIndex        =   21
            Top             =   825
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.OptionButton optArray 
            BackColor       =   &H0050C0A4&
            Caption         =   "Wizards"
            Height          =   495
            Index           =   6
            Left            =   6000
            TabIndex        =   18
            Top             =   400
            Width           =   1455
         End
         Begin VB.CommandButton cmdProceed 
            Caption         =   "Proceed"
            Height          =   375
            Left            =   3000
            TabIndex        =   16
            Top             =   1440
            Width           =   2175
         End
         Begin VB.OptionButton optArray 
            BackColor       =   &H0050C0A4&
            Caption         =   "Administration Tools"
            Height          =   495
            Index           =   5
            Left            =   6000
            TabIndex        =   17
            Top             =   825
            Width           =   1965
         End
         Begin VB.OptionButton optArray 
            BackColor       =   &H0050C0A4&
            Caption         =   "User Group Configuration"
            Height          =   495
            Index           =   2
            Left            =   300
            TabIndex        =   5
            Top             =   825
            Width           =   2175
         End
         Begin VB.OptionButton optArray 
            BackColor       =   &H0050C0A4&
            Caption         =   "User Accounts"
            Height          =   495
            Index           =   1
            Left            =   3120
            TabIndex        =   4
            Top             =   390
            Width           =   1575
         End
         Begin VB.OptionButton optArray 
            BackColor       =   &H0050C0A4&
            Caption         =   "User Groups"
            Height          =   495
            Index           =   0
            Left            =   300
            TabIndex        =   3
            Top             =   400
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame fraConnect 
         BackColor       =   &H0050C0A4&
         Caption         =   "Server Connection"
         ForeColor       =   &H8000000E&
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   2850
         Width           =   8055
         Begin VB.CheckBox chkProxy 
            BackColor       =   &H0050C0A4&
            Caption         =   "Has Proxy"
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox txtProxy 
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   1200
            Width           =   4095
         End
         Begin VB.TextBox txtEncPass 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   6120
            PasswordChar    =   "*"
            TabIndex        =   12
            Top             =   720
            Width           =   1800
         End
         Begin VB.TextBox txtEncPass 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   6120
            PasswordChar    =   "*"
            TabIndex        =   11
            Top             =   480
            Width           =   1800
         End
         Begin VB.CheckBox chkUseEncryption 
            BackColor       =   &H0050C0A4&
            Caption         =   "Encryption used"
            Height          =   255
            Index           =   1
            Left            =   6120
            TabIndex        =   9
            Top             =   240
            Width           =   1545
         End
         Begin VB.ComboBox ComAlgorithm 
            Height          =   315
            Index           =   0
            ItemData        =   "frmDatabaseConnect.frx":1FF27
            Left            =   6960
            List            =   "frmDatabaseConnect.frx":1FF46
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   120
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.TextBox txtServerURL 
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Text            =   "http://www.immap.org/"
            Top             =   540
            Width           =   4095
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "Connect"
            Height          =   315
            Left            =   4440
            TabIndex        =   13
            Top             =   540
            Width           =   915
         End
         Begin VB.Label lblServerURL 
            AutoSize        =   -1  'True
            BackColor       =   &H0050C0A4&
            Caption         =   "Active Server URL:"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   300
            Width           =   1380
         End
         Begin VB.Label lblPass 
            AutoSize        =   -1  'True
            BackColor       =   &H0050C0A4&
            Caption         =   "Pass:"
            Height          =   195
            Index           =   0
            Left            =   5640
            TabIndex        =   10
            Top             =   480
            Width           =   390
         End
         Begin VB.Label lblPass 
            AutoSize        =   -1  'True
            BackColor       =   &H0050C0A4&
            Caption         =   "Key:"
            Height          =   195
            Index           =   2
            Left            =   5640
            TabIndex        =   15
            Top             =   720
            Width           =   315
         End
      End
      Begin VB.TextBox txtAppStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   735
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "frmDatabaseConnect.frx":1FF8A
         Top             =   0
         Width           =   5985
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   0
         Picture         =   "frmDatabaseConnect.frx":1FFB3
         Stretch         =   -1  'True
         ToolTipText     =   "Visit OASIS support website"
         Top             =   0
         Width           =   2310
      End
   End
End
Attribute VB_Name = "frmDatabaseConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents records As ADODB.Recordset
Attribute records.VB_VarHelpID = -1
Private RSUserGroups As New ADODB.Recordset
Private RSUserAccounts As New ADODB.Recordset

Private Declare Function CoCreateGuid _
                Lib "ole32" (id As Any) As Long

Dim WithEvents pGetUsersThread As MThreadVB.Thread
Attribute pGetUsersThread.VB_VarHelpID = -1
Dim WithEvents pLoadUsr As MThreadVB.Thread
Attribute pLoadUsr.VB_VarHelpID = -1
Private m_bServerConnection As Boolean

Dim m_frmUsrAccs As frmUserAccounts
Attribute m_frmUsrAccs.VB_VarHelpID = -1
Dim m_frmUsrGrps As frmUserGroups
Attribute m_frmUsrGrps.VB_VarHelpID = -1
Dim m_frmConfig As frmConfig
Dim m_frmSelectUserGroup As frmSelectUserGroup
Dim m_frmMapPrint As frmMapPrint
Dim m_frmAdminTools As frmAdminTools
Dim m_frmWizards As frmWizards

Public g_bProxyEnabled As Boolean
Public g_sProxy As String

Public Function CreateGUID() As String
        '<EhHeader>
        On Error GoTo CreateGUID_Err
        '</EhHeader>
        Dim id(0 To 15) As Byte
        Dim Cnt As Long, GUID As String

100     If CoCreateGuid(id(0)) = 0 Then

102         For Cnt = 0 To 15
104             CreateGUID = CreateGUID + IIf(id(Cnt) < 16, "0", "") + Hex$(id(Cnt))
106         Next Cnt

108         CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
        Else
110         MsgBox "Error while creating GUID!"
        End If

        '<EhFooter>
        Exit Function

CreateGUID_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.CreateGUID " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub TerminateThreads()
        '<EhHeader>
        On Error GoTo TerminateThreads_Err
        '</EhHeader>
    
100     If Not pGetUsersThread Is Nothing Then
102         If pGetUsersThread.IsThreadRunning Then
104             pGetUsersThread.TerminateWin32Thread
            End If

106         Set pGetUsersThread = Nothing
        End If
    
108     If Not pLoadUsr Is Nothing Then
110         If pLoadUsr.IsThreadRunning Then
112             pLoadUsr.TerminateWin32Thread
            End If

114         Set pLoadUsr = Nothing
        End If
    
        '<EhFooter>
        Exit Sub

TerminateThreads_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.TerminateThreads " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateThreads()
        '<EhHeader>
        On Error GoTo CreateThreads_Err
        '</EhHeader>
    
100     If pGetUsersThread Is Nothing Then
102         Set pGetUsersThread = New Thread
        End If
    
104     If pLoadUsr Is Nothing Then
106         Set pLoadUsr = New Thread
        End If
    
        '<EhFooter>
        Exit Sub

CreateThreads_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.CreateThreads " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CheckIfEncryptedOK() As Boolean
        '<EhHeader>
        On Error GoTo CheckIfEncryptedOK_Err
        '</EhHeader>

        Dim sVariable As String
        Dim oAES As New clsAES
        Dim sReturnValue As String
    
100     CheckIfEncryptedOK = False
        
102     If (Me.txtEncPass(0) = "" Or Me.txtEncPass(1) = "") And chkUseEncryption(1).Value = vbChecked Then

104         MsgBox "Please enter a password and key!"

106     ElseIf chkUseEncryption(1).Value = vbChecked Then
        
108         m_sKey = KeyGen(txtEncPass(1).Text)
            
110         sVariable = "/oasis.asp?sKey=VALIDATE&str=" & txtEncPass(0).Text
112         sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(txtServerURL.Text & sVariable, True)
114         sVariable = oAES.AESEncyptString(txtEncPass(0).Text, m_sKey)

116         If sReturnValue <> sVariable Then
118             CheckIfEncryptedOK = False
            Else
120             CheckIfEncryptedOK = True
            End If
        
122         g_bHasEncrypt = True
        
        Else
124         g_bHasEncrypt = False

126         If m_frmOASISProgress.OpenHttpCommsResponse(WebSite & "oasis.asp?sKey=TEST", True) <> "0" Then
128             CheckIfEncryptedOK = True
            End If
        End If
        
130     Set oAES = Nothing
    
        '<EhFooter>
        Exit Function

CheckIfEncryptedOK_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.CheckIfEncryptedOK " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub chkProxy_Click()
    If chkProxy.Value = vbChecked Then
        txtProxy.Enabled = True
    Else
        txtProxy.Enabled = False
    End If
End Sub

Private Sub chkUseEncryption_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo chkUseEncryption_Click_Err
        '</EhHeader>

100     If chkUseEncryption(1) = vbChecked Then
102         Me.txtEncPass(0).Enabled = True
104         Me.txtEncPass(1).Enabled = True
106         Me.txtEncPass(0).BackColor = vbWhite
108         Me.txtEncPass(1).BackColor = vbWhite
        Else
110         Me.txtEncPass(0).Enabled = False
112         Me.txtEncPass(1).Enabled = False
114         Me.txtEncPass(0).BackColor = &HC0C0C0
116         Me.txtEncPass(1).BackColor = &HC0C0C0
        End If

        '<EhFooter>
        Exit Sub

chkUseEncryption_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.chkUseEncryption_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub cmdConnect_Click()
        '<EhHeader>
        On Error GoTo cmdConnect_Click_Err
        '</EhHeader>

        Dim sString As String
        
100     If Left$(txtServerURL, 7) <> "http://" Then txtServerURL = "http://" & txtServerURL

102     cmdConnect.Enabled = False
104     optOperation.Visible = False
106     Me.Refresh

        If chkProxy.Value = vbChecked Then
            g_bProxyEnabled = True
            g_sProxy = txtProxy.Text
        End If

108     m_bServerConnection = True
110     WebSite = txtServerURL.Text
          
112     If Right$(WebSite, 1) <> "/" Then
114         WebSite = WebSite & "/"
        End If

116     If Not m_frmOASISProgress.OpenHttpCommsResponse(WebSite & "Oasis.asp?doexists=1", True) = "Yes!" Then
        
118         MsgBox "Server could not be found.  Please verify the URL address and also that your computer has internet access."
120         SetStatus "Server could not be found.  Please verify the URL address and also that your computer has internet access."
122         m_frmDebug.DebugPrint "Server could not be found.  Please verify the URL address and also that your computer has internet access."
124         optOperation.Visible = False
126         cmdConnect.Enabled = True
            Exit Sub

        Else
128         SetStatus "Server exists"
130         m_frmDebug.DebugPrint "Server exists"
132         SetStatus getASPFileVersionAndDate(WebSite)

134         SetStatus "Checking if server is encrypted..."
136         m_frmDebug.DebugPrint "Checking if server is encrypted..."

138         If chkUseEncryption(1).Value = vbChecked Then
140             If Not CheckIfEncryptedOK Then
142                 MsgBox "It seems like the server is encrypted," & vbCrLf & "You will need to provide the correct Encryption details before connecting..."
144                 optOperation.Visible = False
146                 cmdConnect.Enabled = True
                    Exit Sub
                End If

148             SetStatus "Encrypted connection successful..."
150             m_frmDebug.DebugPrint "Encrypted connection successful..."

            Else

152             If CheckIfEncryptedOK Then
154                 MsgBox "It seems like the server is encrypted," & vbCrLf & "You will need to provide the correct Encryption details before connecting..."
156                 optOperation.Visible = False
158                 cmdConnect.Enabled = True
                    Exit Sub
                Else
160                 SetStatus "Unencrypted connection successful..."
162                 m_frmDebug.DebugPrint "Unencrypted connection successful..."

                End If

            End If

164         SetStatus "Connecting to Server..."
166         m_frmDebug.DebugPrint "Connecting to Server..."

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'GET USER GROUPS
168         SetStatus "Getting User Groups From... " & txtServerURL.Text

170         If RSUserGroups.State = adStateOpen Then
172             RSUserGroups.Close
174             Set RSUserGroups = Nothing
            End If

176         Set RSUserGroups = New ADODB.Recordset
178         sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM UserGroups ORDER BY Name")
180         Set RSUserGroups = m_frmOASISProgress.OpenHttpCommsRS(sString, True)

182         If Not RSUserGroups.State = adStateClosed Then

184             If RSUserGroups.EOF And RSUserGroups.Bof Then
186                 SetStatus "hmmmm.... User Group EOF and BOF where true..."
188                 Call disableAdvancedFeatures(False)
190                 MsgBox "No user groups found! You have to create at least 1 (one) User Group before you can continue configuration. Refer to OASIS Admin Handbook how to enter data.", vbInformation, "OASIS Admin Tool"
                Else
192                 RSUserGroups.MoveFirst
194                 Call disableAdvancedFeatures(True)
196                 SetStatus "Success!"
                End If

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'GET USER ACCOUNTS
198             SetStatus "Getting User Accounts From... " & txtServerURL.Text

200             Set RSUserAccounts = New ADODB.Recordset

202             sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM Users")
204             m_frmDebug.DebugPrint sString
206             Set RSUserAccounts = m_frmOASISProgress.OpenHttpCommsRS(sString, True)
               
208             If Not RSUserAccounts.State = adStateClosed Then

210                 If RSUserAccounts.EOF And RSUserAccounts.Bof Then
212                     SetStatus "hmmmm.... User Group EOF and BOF where true..."

214                     MsgBox "No user accounts found! You have to create at least 1 (one) User Account users can use the OASIS Client online. Refer to OASIS Admin Handbook how to enter data.", vbInformation, "OASIS Admin Tool"
                    Else
216                     RSUserAccounts.MoveFirst

218                     SetStatus "Success!"

                    End If
                    
                End If

            Else
            
220             MsgBox "User groups failed to load"
            
            End If

222         optOperation.Visible = True
        End If

224     m_bServerConnection = False
226     cmdConnect.Enabled = True

        '<EhFooter>
        Exit Sub

cmdConnect_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.cmdConnect_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub disableAdvancedFeatures(bSetting As Boolean)
        '<EhHeader>
        On Error GoTo disableAdvancedFeatures_Err
        '</EhHeader>

100     Me.optArray(1).Visible = bSetting
102     Me.optArray(2).Visible = bSetting
108     Me.optArray(5).Visible = bSetting
110     Me.optArray(6).Visible = bSetting
        'Me.optArray(3).Visible = bSetting

        '<EhFooter>
        Exit Sub

disableAdvancedFeatures_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.disableAdvancedFeatures " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub doProgress(bStart As Boolean, _
                       Optional sText As String, _
                       Optional lProc As Long)
        '<EhHeader>
        On Error GoTo doProgress_Err
        '</EhHeader>

        On Error Resume Next
    
100     If bStart Then
102         SetStatus "(begin progress...)"
        Else
104         SetStatus "(end Progress)"
        End If
    
106     If bStart Then
108         frmProgress.Show vbModeless, Me

110         With frmProgress
    
112             .Timer1.Enabled = bStart
114             .dxProgressBar.Visible = bStart
    
            End With
    
        Else
116         frmProgress.Hide
118         Unload frmProgress
        End If
        
        '<EhFooter>
        Exit Sub

doProgress_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.doProgress " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdProceed_Click()
        '<EhHeader>
        On Error GoTo cmdProceed_Click_Err
        '</EhHeader>
        
100     If optArray(0).Value = True Then

102         m_frmUsrGrps.setUserGroupsRS RSUserGroups
104         m_frmUsrGrps.setUserAccountsRS RSUserAccounts
106         m_frmUsrGrps.Show vbModeless, Me

108     ElseIf optArray(3).Value = True Then

            MsgBox "Please use MSSQL Server Managment Studio instead of this feature"
110         'frmDatabaseExplorer.Show vbModeless, Me

112     ElseIf optArray(1).Value = True Then

114         Me.Enabled = False
116         m_frmUsrAccs.dxUsers.Columns.ColumnByFieldName("UserGroupID").ColumnType = gedLookupEdit

118         With m_frmUsrAccs.dxUsers.Columns.ColumnByFieldName("UserGroupID").LookupColumn

120             .LookupDatasetType = dtADODataset
122             Set .DataSource = RSUserGroups
124             .LookupDataset.open
126             .LookupKeyField = "ID"
128             .LookupResultField = "Name"
130             .ListFieldName = "Name"

            End With

132         With RSUserGroups

134             If SafeMoveFirst(RSUserGroups) Then
            
136                 Do While Not .EOF
138                     m_frmUsrAccs.cmbUserGroups.AddItem RSUserGroups.fields.Item("Name")
140                     m_frmUsrAccs.cmbUserGroups.ItemData(m_frmUsrAccs.cmbUserGroups.ListCount - 1) = RSUserGroups.fields.Item("ID")
142                     .MoveNext
                    Loop

                End If

144             SafeMoveFirst RSUserGroups

            End With

146         Call m_frmUsrAccs.setUserAccountsRS(RSUserAccounts)
148         Set m_frmUsrAccs.dxUsers.DataSource = RSUserAccounts
150         m_frmUsrAccs.dxUsers.Columns.RetrieveFields
152         m_frmUsrAccs.Show vbModeless, Me
154         Me.Enabled = True

156     ElseIf optArray(2).Value = True Then

158         If Not RSUserGroups Is Nothing Then
160             Set m_frmSelectUserGroup.dxUG.DataSource = Nothing
162             Set m_frmSelectUserGroup.dxUG.DataSource = RSUserGroups
164             m_frmSelectUserGroup.dxUG.Columns.RetrieveFields
166             m_frmSelectUserGroup.dxUG.Dataset.Refresh
            End If
            
168         m_frmSelectUserGroup.Show vbModal

170         If m_frmSelectUserGroup.Tag = True Then
172             m_frmConfig.SetWebsite WebSite
174             m_frmConfig.SetTablePrefix RSUserGroups.fields("Name").Value

176             If m_frmConfig.LoadUserData("") Then m_frmConfig.Show vbModeless, Me
            End If

178     ElseIf optArray(5).Value = True Then

180         m_frmAdminTools.Show vbModeless, Me

182     ElseIf optArray(6).Value = True Then

184         m_frmWizards.setUserGroupsRS RSUserGroups
186         m_frmWizards.Show vbModeless, Me
        
        End If

        '<EhFooter>
        Exit Sub

cmdProceed_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.cmdProceed_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = vbKeyF11 Then
        ShowImmediatePane
    End If

End Sub

Private Sub txtServerURL_KeyUp(KeyCode As Integer, _
                               Shift As Integer)
        '<EhHeader>
        On Error GoTo txtServerURL_KeyUp_Err
        '</EhHeader>
    
100     If KeyCode = vbKeyReturn Then cmdConnect_Click
    
        '<EhFooter>
        Exit Sub

txtServerURL_KeyUp_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.txtServerURL_KeyUp " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function ShowImmediatePane() As Boolean
        '<EhHeader>
        On Error GoTo ShowImmediatePane_Err
        '</EhHeader>

100     m_frmDebug.Show
102     ShowImmediatePane = True

        '<EhFooter>
        Exit Function

ShowImmediatePane_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.ShowImmediatePane " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub Form_Initialize()
        '<EhHeader>
        On Error GoTo Form_Initialize_Err
        '</EhHeader>

100     If App.PrevInstance = True Then
102         MsgBox "App Already Running. You cannot run two Instances of OASIS Server Admin. ", vbOKOnly + vbCritical + vbApplicationModal
104         End
        End If

106     Set m_frmDebug = New frmDebug
108     Debug.Assert ShowImmediatePane
        
110     Set g_PictureDialogLarge = Me.PictureDialogLarge.Picture
112     Set g_PictureDialogSmall = Me.PictureDialogSmall.Picture
                
        '<EhFooter>
        Exit Sub

Form_Initialize_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.Form_Initialize " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ResetStuff()
        '<EhHeader>
        On Error GoTo ResetStuff_Err
        '</EhHeader>

100     txtServerURL.Text = GetSetting(App.EXEName, "Settings", "WebServer", "http://www.immap.org/")
        '<EhFooter>
        Exit Sub

ResetStuff_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.ResetStuff " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

        Dim mFileSysObj As New FileSystemObject
100     Set m_frmUsrGrps = New frmUserGroups
102     Set m_frmUsrAccs = New frmUserAccounts
104     Set m_frmConfig = New frmConfig
106     Set m_frmSelectUserGroup = New frmSelectUserGroup
108     Set m_frmAdminTools = New frmAdminTools
110     Set m_frmWizards = New frmWizards

112     Set m_frmOASISProgress = New frmOASISProgress
114     m_frmOASISProgress.InitialiseProgressForm CreateAppPath & "\data\db\Oasisclient.mdb"
116     m_frmOASISProgress.SetShowAdvanced False
    
118     Me.Caption = "OASIS Server Administration Toolbox. Version: " & App.major & "." & App.minor & "." & App.Revision & " Developed by iMMAP.org Contact: support@iMMAP.org © Copyright iMMAP.org 2003-2008"
    
120     CreateThreads
122     ResetStuff
124     Set mFileSysObj = Nothing

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    Dim sLog As String
    Dim i As Integer
    
    On Error Resume Next
        
    If Not m_frmOASISProgress Is Nothing Then
        If m_frmOASISProgress.Visible Then
            MsgBox "Please wait until the server connection completes!", vbCritical, "Communication in progress"
            Cancel = True
            m_frmOASISProgress.SetFocus
            Exit Sub
        End If
    End If

    If Not WebSite = "" Then
    
        SetStatus "Writing log file..."
        sLog = Replace(Me.txtAppStatus.Text, vbCrLf, " -- ")
        sLog = Replace(sLog, vbCr, " -- ")
        sLog = Replace(sLog, vbLf, " -- ")
        sLog = Replace(sLog, vbLf, " -- ")
        m_frmDebug.DebugPrint WebSite & "Oasis.asp?droplog=" & CheckEncrypt(sLog)

        If m_frmOASISProgress.OpenHttpCommsResponse(WebSite & "Oasis.asp?droplog=" & CheckEncrypt(Left(sLog, 255)), True) = "1" Then
            m_frmDebug.DebugPrint "Log file write successful"
        Else
            m_frmDebug.DebugPrint "Log file write UNsuccessful"
        End If
    End If

    TerminateThreads
    SaveSetting App.EXEName, "Settings", "WebServer", txtServerURL.Text
            
    On Error Resume Next
    records.Close
    Set records = Nothing
    Set m_oAES = Nothing

    Set g_PictureDialogSmall = Nothing
    Set g_PictureDialogLarge = Nothing
    
    RSUserAccounts.Close
    Set RSUserAccounts = Nothing
    
    RSUserGroups.Close
    
    If Not m_frmDebug Is Nothing Then
        Unload m_frmDebug
        Set m_frmDebug = Nothing
    End If
    
    Set RSUserGroups = Nothing

    Set m_frmUsrGrps = Nothing
    Set m_frmUsrAccs = Nothing
    Set m_frmConfig = Nothing
    Set m_frmSelectUserGroup = Nothing
    Set m_frmAdminTools = Nothing
    Set m_frmWizards = Nothing
    Set m_frmOASISProgress = Nothing

    For i = 1 To Forms.Count

        If Forms(i).Name <> Me.Name Then
            Unload Forms(i)
            Set Forms(i) = Nothing
        End If

    Next

End Sub

Private Sub Image1_Click()
        '<EhHeader>
        On Error GoTo Image1_Click_Err
        '</EhHeader>
100     ShellExecute Me.hwnd, vbNullString, "http://oasis.comindwork.com", vbNullString, vbNullString, 1
        '<EhFooter>
        Exit Sub

Image1_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmDatabaseConnect.Image1_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

