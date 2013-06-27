VERSION 5.00
Begin VB.Form frmEncryptWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Encryption Wizard"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmEncryptWizard.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTools 
      Caption         =   "Admin Tools"
      Height          =   375
      Left            =   4800
      TabIndex        =   25
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame FraServerSettings 
      BackColor       =   &H0050C0A4&
      Caption         =   "Choose operation:"
      Height          =   780
      Index           =   3
      Left            =   1920
      TabIndex        =   22
      Top             =   840
      Width           =   4155
      Begin VB.OptionButton optOperation 
         BackColor       =   &H0050C0A4&
         Caption         =   "Decrypt Server"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optOperation 
         BackColor       =   &H0050C0A4&
         Caption         =   "Encrypt Server"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame FraServerSettings 
      BackColor       =   &H0050C0A4&
      Caption         =   "Key generation:"
      Height          =   660
      Index           =   2
      Left            =   0
      TabIndex        =   17
      Top             =   4200
      Width           =   6075
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Test"
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   20
         Top             =   3120
         Width           =   525
      End
      Begin VB.TextBox txtKeyFile 
         Height          =   315
         Index           =   0
         Left            =   900
         TabIndex        =   19
         Top             =   210
         Width           =   4425
      End
      Begin VB.CommandButton cmdLocalPath 
         Caption         =   "..."
         Height          =   285
         Left            =   5370
         TabIndex        =   18
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackColor       =   &H0050C0A4&
         Caption         =   "Key File:"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   21
         Top             =   330
         Width           =   600
      End
   End
   Begin VB.Frame FraServerSettings 
      BackColor       =   &H0050C0A4&
      Caption         =   "Encryption Settings:"
      Height          =   1620
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   2520
      Width           =   6075
      Begin VB.ComboBox ComAlgorithm 
         Height          =   315
         Index           =   1
         ItemData        =   "frmEncryptWizard.frx":6852
         Left            =   1320
         List            =   "frmEncryptWizard.frx":6871
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   4635
      End
      Begin VB.TextBox txtEncPass 
         ForeColor       =   &H000000FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   720
         Width           =   1665
      End
      Begin VB.TextBox txtEncPass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   4230
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   750
         Width           =   1665
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Test"
         Height          =   315
         Index           =   4
         Left            =   60
         TabIndex        =   8
         Top             =   3120
         Width           =   525
      End
      Begin VB.TextBox txtEncPass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   4230
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1110
         Width           =   1665
      End
      Begin VB.TextBox txtEncPass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1140
         Width           =   1665
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackColor       =   &H0050C0A4&
         Caption         =   "Type:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackColor       =   &H0050C0A4&
         Caption         =   "Password:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         BackColor       =   &H0050C0A4&
         Caption         =   "Key:"
         Height          =   195
         Index           =   1
         Left            =   3210
         TabIndex        =   14
         Top             =   810
         Width           =   1035
      End
      Begin VB.Label lblPass 
         BackColor       =   &H0050C0A4&
         Caption         =   "Type Again:"
         Height          =   315
         Index           =   10
         Left            =   240
         TabIndex        =   13
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label lblPass 
         BackColor       =   &H0050C0A4&
         Caption         =   "Type Again:"
         Height          =   285
         Index           =   0
         Left            =   3180
         TabIndex        =   12
         Top             =   1140
         Width           =   1065
      End
   End
   Begin VB.Frame FraServerSettings 
      BackColor       =   &H0050C0A4&
      Caption         =   "Specify the OASIS Server:"
      Height          =   780
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   6075
      Begin VB.TextBox txtServerURL 
         Height          =   330
         Left            =   510
         TabIndex        =   3
         Top             =   300
         Width           =   5385
      End
      Begin VB.Label lblServerURL 
         AutoSize        =   -1  'True
         BackColor       =   &H0050C0A4&
         Caption         =   "URL:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   390
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5100
      TabIndex        =   1
      Top             =   4920
      Width           =   885
   End
   Begin VB.CommandButton cmdCommit 
      Caption         =   "Save"
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   4920
      Width           =   885
   End
End
Attribute VB_Name = "frmEncryptWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub GenerateEncryptionKeyFile(SourceFileName As String, _
                                     NrOfKeys As Long)
        '<EhHeader>
        On Error GoTo GenerateEncryptionKeyFile_Err
        '</EhHeader>
        Dim FileNr1 As Integer
        Dim i As Long
        Dim sString As String
        Dim ky As Long
    
100     If Len(SourceFileName) = 0 Then Exit Sub
    
102     FileNr1 = FreeFile
104     Open SourceFileName For Output Shared As #FileNr1
106     Randomize

108     For i = 1 To NrOfKeys
110         Print #FileNr1, Int((ky * Rnd) + 1)
112     Next i
    
114     Close #FileNr1

        '<EhFooter>
        Exit Sub

GenerateEncryptionKeyFile_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmEncryptWizard.GenerateEncryptionKeyFile " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetEncPWD(sPWD As String, _
                           sKey As String, _
                           sEncType As String) As String
        '<EhHeader>
        On Error GoTo GetEncPWD_Err
        '</EhHeader>
    
        Dim oRC4 As clsRC4
        Dim oAES As clsAES
        Dim oDES As clsDES
        Dim oBlowFish As clsBlowfish
        Dim oGost As clsGost
        Dim oSerpent As clsSerpent
        Dim oSkipjack As clsSkipjack
        Dim oTEA As clsTEA
        Dim oTwoFish As clsTwofish

100     Select Case sEncType

            Case "RC4"
                'OK
102             Set oRC4 = New clsRC4
104             GetEncPWD = oRC4.RC4(sPWD, sKey)
        
106             Set oRC4 = Nothing
            
108         Case "AES"
                'OK
110             Set oAES = New clsAES
            
112             GetEncPWD = oAES.AESEncyptString(sPWD, sKey)
            
114             Set oAES = Nothing
            
116         Case "DES"
118             Set oDES = New clsDES
            
120             GetEncPWD = oDES.EncryptString(sPWD, sKey)
            
122             Set oDES = Nothing
            
124         Case "BlowFish"
126             Set oBlowFish = New clsBlowfish

128             GetEncPWD = oBlowFish.EncryptString(sPWD, sKey)
            
130             Set oBlowFish = Nothing

132         Case "Gost"
134             Set oGost = New clsGost

136             GetEncPWD = oGost.EncryptString(sPWD, sKey)
            
138             Set oGost = Nothing

140         Case "Serpent"
142             Set oSerpent = New clsSerpent
144             GetEncPWD = oSerpent.EncryptString(sPWD, sKey)
            
146             Set oSerpent = Nothing

148         Case "Skipjack"
150             Set oSkipjack = New clsSkipjack

152             GetEncPWD = oSkipjack.EncryptString(sPWD, sKey)
154             Set oSkipjack = Nothing

156         Case "TEA"
158             Set oTEA = New clsTEA

160             GetEncPWD = oTEA.EncryptString(sPWD, sKey)
            
162             Set oTEA = Nothing

164         Case "Twofish"
                        
166             Set oTwoFish = New clsTwofish
168             GetEncPWD = oTwoFish.EncryptString(sPWD, sKey)
            
170             Set oTwoFish = Nothing
        End Select

        '<EhFooter>
        Exit Function

GetEncPWD_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmEncryptWizard.GetEncPWD " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmEncryptWizard.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Connect(Index As Integer)
        '<EhHeader>
        On Error GoTo Connect_Err
        '</EhHeader>

        Dim sVariable As String
        Dim fs As New FileSystemObject
        Dim oFile As Object
        Dim sKey As String
        Dim i As Integer
        Dim oRD4 As New clsRC4
        Dim oAES As New clsAES
        Dim sReturnValue As String
    
100     Select Case Index
    
            Case 1
            
102             sKey = KeyGen(txtEncPass(2).Text)
104             sVariable = "/oasis.asp?sKey=CREATE&str=" & sKey & "&dummyvar=" & Now()
106             sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(txtServerURL.Text & sVariable, True)
            
108             m_frmDebug.DebugPrint txtServerURL.Text & sVariable
            
110             If (fs.FileExists(txtKeyFile(0).Text)) = True Then
112                 fs.DeleteFile txtKeyFile(0).Text, True
                End If

114             Set oFile = fs.CreateTextFile(txtKeyFile(0).Text)
116             oFile.WriteLine sKey & vbCrLf & oAES.AESEncyptString(txtEncPass(3).Text, sKey)
118             oFile.Close
            
120             If sReturnValue = "2" Then
122                 MsgBox "Server encrypted"
124                 g_bHasEncrypt = True
126                 frmDatabaseConnect.chkUseEncryption(1).Value = vbChecked
128                 frmDatabaseConnect.txtEncPass(0).Text = txtEncPass(3).Text
130                 frmDatabaseConnect.txtEncPass(1).Text = txtEncPass(2).Text
                
132             ElseIf sReturnValue = "1" Then
134                 MsgBox "Server is already encrypted"
                Else
136                 MsgBox "Server encryption unsuccessful (error: " & sReturnValue & ")"
                End If

138         Case 2
            
140             sKey = KeyGen(txtEncPass(2).Text)
142             sVariable = "/oasis.asp?sKey=DELETEKEY&str=" & sKey & "&dummyvar=" & Now()
144             sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(txtServerURL.Text & sVariable, True)

146             m_frmDebug.DebugPrint txtServerURL.Text & sVariable

148             If sReturnValue = "1" Then
150                 MsgBox "Server decryption successful"
152                 g_bHasEncrypt = False
154                 frmDatabaseConnect.chkUseEncryption(1).Value = vbUnchecked
                
156             ElseIf sReturnValue = "2" Then
158                 MsgBox "Server is already decrypted"
                Else
160                 MsgBox "Server encryption unsuccessful (error: " & sReturnValue & ")"
                End If
                       
        End Select

162     DoEvents

        '<EhFooter>
        Exit Sub

Connect_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmEncryptWizard.Connect " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function URLEncode(sRawURL As String) As String
        '<EhHeader>
        On Error GoTo URLEncode_Err
        '</EhHeader>

        On Error GoTo Catch
        Dim iLoop As Integer
        Dim sRtn As String
        Dim sTmp As String
        Const sValidChars = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz:/.?=_-$(){}~&"

100     If Len(sRawURL) > 0 Then
            ' Loop through each char

102         For iLoop = 1 To Len(sRawURL)
104             sTmp = Mid(sRawURL, iLoop, 1)

106             If InStr(1, sValidChars, sTmp, vbBinaryCompare) = 0 Then
                    ' If not ValidChar, convert to HEX and p
                    '     refix with %
108                 sTmp = Hex(Asc(sTmp))

110                 If sTmp = "20" Then
112                     sTmp = "+"
114                 ElseIf Len(sTmp) = 1 Then
116                     sTmp = "%0" & sTmp
                    Else
118                     sTmp = "%" & sTmp
                    End If

                End If

120             sRtn = sRtn & sTmp
122         Next iLoop

124         URLEncode = sRtn
        End If

Finally:
        Exit Function
Catch:
126     URLEncode = ""
128     Resume Finally
        '<EhFooter>
        Exit Function

URLEncode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmEncryptWizard.URLEncode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function



Private Sub cmdLocalPath_Click()
        '<EhHeader>
        On Error GoTo cmdLocalPath_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog
    
        On Error Resume Next
100     c.DefaultExt = "OKEY"
102     c.Filter = "*.OKEY"
104     c.ShowSave
106     txtKeyFile(0).Text = c.fileName
        '<EhFooter>
        Exit Sub

cmdLocalPath_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmEncryptWizard.cmdLocalPath_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCommit_Click()
        '<EhHeader>
        On Error GoTo cmdCommit_Click_Err
        '</EhHeader>
        
100     Select Case optOperation(0)
    
            Case vbChecked

102             If Len(txtServerURL.Text) < 5 Then
104                 MsgBox "It seems like the server URL is incorrect:" & txtServerURL.Text
106             ElseIf Len(txtEncPass(3).Text) < 6 Then
108                 MsgBox "It seems like the Password is too short."
110             ElseIf txtEncPass(3).Text <> txtEncPass(6).Text Then
112                 MsgBox "The passwords do not match.  Please retry."
114             ElseIf txtEncPass(2).Text <> txtEncPass(7).Text Then
116                 MsgBox "The keys do not match.  Please retry."
118             ElseIf Len(txtEncPass(2).Text) < 6 Then
120                 MsgBox "It seems like the Key is too short."
122             ElseIf IsNull(Me.txtKeyFile(0).Text) Or txtKeyFile(0).Text = "" Then
124                 MsgBox "It seems like the key file path is incorrect"
                Else
126                 Connect 1
                End If

128         Case Else
        
130             If Len(txtServerURL.Text) < 5 Then
132                 MsgBox "It seems like the server URL is incorrect:" & txtServerURL.Text
134             ElseIf Len(txtEncPass(3).Text) < 6 Then
136                 MsgBox "It seems like the Password is too short."
138             ElseIf Len(txtEncPass(2).Text) < 6 Then
140                 MsgBox "It seems like the Key is too short."
                Else
142                 Connect 2
                End If
    
        End Select
    
        '<EhFooter>
        Exit Sub

cmdCommit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmEncryptWizard.cmdCommit_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTools_Click()
        '<EhHeader>
        On Error GoTo cmdTools_Click_Err
        '</EhHeader>
        Dim m_frmAdminTools As frmAdminTools
100     Set m_frmAdminTools = New frmAdminTools
102     m_frmAdminTools.Show vbModeless, Me
104     Set m_frmAdminTools = Nothing
        '<EhFooter>
        Exit Sub

cmdTools_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmEncryptWizard.cmdTools_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
        Dim i As Integer

        On Error Resume Next
100     ComAlgorithm(0).ListIndex = 0
102     ComAlgorithm(1).ListIndex = 0
104     ComAlgorithm(2).ListIndex = 0

106     txtServerURL.Text = GetSetting(App.EXEName, "Settings", "WebServer", "http://www.immap.org/")
108     Me.Picture = g_PictureDialogSmall
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmEncryptWizard.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub optOperation_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo optOperation_Click_Err
        '</EhHeader>

100     Select Case optOperation(0)
    
            Case vbChecked

102             lblPass(0).Visible = True
104             lblPass(10).Visible = True
106             txtEncPass(6).Visible = True
108             txtEncPass(7).Visible = True
110             Me.FraServerSettings(2).Visible = True

112         Case Else
        
114             lblPass(0).Visible = False
116             lblPass(10).Visible = False
118             txtEncPass(6).Visible = False
120             txtEncPass(7).Visible = False
122             Me.FraServerSettings(2).Visible = False
    
        End Select

        '<EhFooter>
        Exit Sub

optOperation_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmEncryptWizard.optOperation_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtEncPass_Change(Index As Integer)
        '<EhHeader>
        On Error GoTo txtEncPass_Change_Err
        '</EhHeader>

100     Select Case Index
    
            Case 3, 2

102             If Len(txtEncPass(Index).Text) < 6 Then
104                 txtEncPass(Index).ForeColor = vbRed
                Else
106                 txtEncPass(Index).ForeColor = vbBlack
                End If
            
        End Select

        '<EhFooter>
        Exit Sub

txtEncPass_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmEncryptWizard.txtEncPass_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
