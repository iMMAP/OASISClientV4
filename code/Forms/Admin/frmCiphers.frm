VERSION 5.00
Begin VB.Form frmCiphers 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0050C0A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Cryptography Tool "
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5280
   Icon            =   "frmCiphers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraEncryptionKey 
      BackColor       =   &H0050C0A4&
      Caption         =   "Encryption Key Password:"
      Height          =   585
      Left            =   150
      TabIndex        =   22
      Top             =   5340
      Width           =   3315
      Begin VB.TextBox txtOASISPWD1234 
         Height          =   285
         Left            =   60
         TabIndex        =   23
         Text            =   "OASISPWD1234"
         Top             =   240
         Width           =   3165
      End
   End
   Begin VB.CommandButton cmdGenerateKey 
      Caption         =   "Generate Key"
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   5490
      Width           =   1515
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   3510
      TabIndex        =   20
      Top             =   1920
      Width           =   1515
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   3450
      TabIndex        =   19
      Top             =   3630
      Width           =   1515
   End
   Begin VB.CommandButton cmdHashAlg 
      Caption         =   "Hash Alg."
      Height          =   375
      Left            =   210
      TabIndex        =   18
      Top             =   4950
      Width           =   1515
   End
   Begin VB.CommandButton cmdEncryptFile 
      Caption         =   "Encrypt File"
      Height          =   375
      Left            =   1860
      TabIndex        =   17
      Top             =   4950
      Width           =   1515
   End
   Begin VB.CommandButton cmdDecryptFile 
      Caption         =   "Decrypt File"
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   4980
      Width           =   1515
   End
   Begin VB.TextBox txtSaltD 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      TabIndex        =   6
      Text            =   "12345"
      Top             =   3660
      Width           =   1215
   End
   Begin VB.ComboBox cmbAlgorithms 
      Height          =   315
      ItemData        =   "frmCiphers.frx":6852
      Left            =   240
      List            =   "frmCiphers.frx":6880
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   4695
   End
   Begin VB.TextBox txtSaltE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      TabIndex        =   3
      Text            =   "12345"
      Top             =   1980
      Width           =   1215
   End
   Begin VB.TextBox txtKeyD 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      TabIndex        =   5
      Text            =   "secretOASISkey"
      Top             =   3660
      Width           =   1815
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmCiphers.frx":690B
      Top             =   1020
      Width           =   4695
   End
   Begin VB.TextBox txtText2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4260
      Width           =   4695
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2700
      Width           =   4695
   End
   Begin VB.TextBox txtKeyE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      TabIndex        =   2
      Text            =   "secretOASISkey"
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Salt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   3420
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Algorithm"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Salt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Key to Decrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3420
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Text to Encrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   780
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Decrypted Text"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4020
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypted String (Base64 Format)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2460
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Key to Encrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1740
      Width           =   1815
   End
End
Attribute VB_Name = "frmCiphers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dude As Integer

Private Sub cmbAlgorithms_Click()
    txtOutput.Text = ""
End Sub

Private Sub cmdDecrypt_Click()
        '<EhHeader>
        On Error GoTo cmdDecrypt_Click_Err
        '</EhHeader>
100     txtText2.Text = DecryptString(cmbAlgorithms.ListIndex, txtOutput, True, txtKeyD, txtSaltD)
        '<EhFooter>
        Exit Sub

cmdDecrypt_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_Cryptography_Tool.frmCiphers.cmdDecrypt_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdDecryptFile_Click()
        '<EhHeader>
        On Error GoTo cmdDecryptFile_Click_Err
        '</EhHeader>
        Dim X As Boolean, Key As String, Salt As String, File1 As String, File2 As String
    
100     File1 = GetFileInName("File To Encrypt/Decrypt", "*.*|*.*")

102     If File1 = "" Then Exit Sub
    
104     File2 = GetFileOutName("Save Encrypted/Decrypted File As", "*.*|*.*")

106     If File2 = "" Then Exit Sub
    
108     Key = InputBox("Enter key:", "Utilize ebCrypt")
110     Salt = InputBox("Enter salt:", "Utilize ebCrypt")
112     X = DecryptFile(cmbAlgorithms.ListIndex, File1, File2, True, True, Key, Salt)

        '<EhFooter>
        Exit Sub

cmdDecryptFile_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_Cryptography_Tool.frmCiphers.cmdDecryptFile_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdEncrypt_Click()
        '<EhHeader>
        On Error GoTo cmdEncrypt_Click_Err
        '</EhHeader>
100     txtOutput.Text = EncryptString(cmbAlgorithms.ListIndex, txtText, True, txtKeyE, txtSaltE)

        '<EhFooter>
        Exit Sub

cmdEncrypt_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_Cryptography_Tool.frmCiphers.cmdEncrypt_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdEncryptFile_Click()
        '<EhHeader>
        On Error GoTo cmdEncryptFile_Click_Err
        '</EhHeader>
        Dim X As Boolean, Key As String, Salt As String, File1 As String, File2 As String
       
100     File1 = GetFileInName("File To Encrypt/Decrypt", "*.*|*.*")

102     If File1 = "" Then Exit Sub
    
104     File2 = GetFileOutName("Save Encrypted/Decrypted File As", "*.*|*.*")

106     If File2 = "" Then Exit Sub
    
108     Key = InputBox("Enter key:", "Utilize ebCrypt")
110     Salt = InputBox("Enter salt:", "Utilize ebCrypt")
112     X = EncryptFile(cmbAlgorithms.ListIndex, File1, File2, True, True, Key, Salt)

        '<EhFooter>
        Exit Sub

cmdEncryptFile_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_Cryptography_Tool.frmCiphers.cmdEncryptFile_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdGenerateKey_Click()
        '<EhHeader>
        On Error GoTo cmdGenerateKey_Click_Err
        '</EhHeader>
100     GenerateEncryptionKeyFile GetFileOutName("Save Key File As", "*.Key|*.Key"), CLng(InputBox("Enter The Length Of desired key length", "OASIS Key Length", "130000"))
        '<EhFooter>
        Exit Sub

cmdGenerateKey_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_Cryptography_Tool.frmCiphers.cmdGenerateKey_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

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
    
        '    On Error GoTo ErrorTrap
102     FileNr1 = FreeFile
104     Open SourceFileName For Output Shared As #FileNr1
106     Randomize

108     For i = 1 To NrOfKeys
110         Print #FileNr1, Int((ky * Rnd) + 1)
112     Next i
    
114     Close #FileNr1
        Exit Sub
        'ErrorTrap:
116     MsgBox "Error " & Err.Number & " " & Err.Description
        '<EhFooter>
        Exit Sub

GenerateEncryptionKeyFile_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_Cryptography_Tool.frmCiphers.GenerateEncryptionKeyFile " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdHashAlg_Click()
        '<EhHeader>
        On Error GoTo cmdHashAlg_Click_Err
        '</EhHeader>
100     frmCiphers.Hide
102     frmHASH.Show vbModal
104     frmCiphers.Show

        '<EhFooter>
        Exit Sub

cmdHashAlg_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_Cryptography_Tool.frmCiphers.cmdHashAlg_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     cmbAlgorithms.ListIndex = 0
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_Cryptography_Tool.frmCiphers.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

