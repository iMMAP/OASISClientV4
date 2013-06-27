VERSION 5.00
Begin VB.Form frmHASH 
   BackColor       =   &H0050C0A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS HASH Algorithm Generation Tool"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5280
   Icon            =   "frmHASH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHash 
      Caption         =   "Generate Hash"
      Height          =   375
      Left            =   3660
      TabIndex        =   6
      Top             =   1830
      Width           =   1245
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
      TabIndex        =   2
      Top             =   2490
      Width           =   4695
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
      Text            =   "frmHASH.frx":6852
      Top             =   1050
      Width           =   4695
   End
   Begin VB.ComboBox cmbAlgorithms 
      Height          =   315
      ItemData        =   "frmHASH.frx":686A
      Left            =   240
      List            =   "frmHASH.frx":687A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   330
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "HASH String (HEX Format)"
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
      TabIndex        =   5
      Top             =   2250
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Text to HASH"
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
      TabIndex        =   4
      Top             =   810
      Width           =   1935
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
      TabIndex        =   3
      Top             =   90
      Width           =   1575
   End
End
Attribute VB_Name = "frmHASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHash_Click()
        '<EhHeader>
        On Error GoTo cmdHash_Click_Err
        '</EhHeader>
100     txtOutput = Hash(cmbAlgorithms.ListIndex, txtText.Text)
        '<EhFooter>
        Exit Sub

cmdHash_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmHASH.cmdHash_Click " & _
               "at line " & Erl
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
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmHASH.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


