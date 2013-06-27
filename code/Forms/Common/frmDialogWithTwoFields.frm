VERSION 5.00
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Begin VB.Form frmDialogWithTwoFields 
   BackColor       =   &H0050C0A4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   2235
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmDialogWithTwoFields.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XpressEditorsLibCtl.dxMemoEdit txt2 
      Height          =   375
      Left            =   180
      OleObjectBlob   =   "frmDialogWithTwoFields.frx":6852
      TabIndex        =   5
      Top             =   1800
      Width           =   5685
   End
   Begin VB.TextBox txt1 
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   570
      Width           =   4335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl2 
      BackColor       =   &H0050C0A4&
      BackStyle       =   0  'Transparent
      Caption         =   "This is label 1"
      Height          =   765
      Left            =   210
      TabIndex        =   2
      Top             =   1080
      Width           =   5775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl1 
      BackColor       =   &H0050C0A4&
      BackStyle       =   0  'Transparent
      Caption         =   "This is label 1"
      Height          =   495
      Left            =   180
      TabIndex        =   1
      Top             =   90
      Width           =   4455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDialogWithTwoFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bClickedOK As Boolean
Public sText1 As String
Public sText2 As String

Private Sub CancelButton_Click()
    bClickedOK = False
    Unload Me
End Sub

Private Sub Form_Load()
    bClickedOK = False
End Sub

Private Sub OKButton_Click()
    bClickedOK = True
    sText1 = txt1
    sText2 = txt2
    Unload Me
End Sub

Private Sub txt1_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = 13 Then
        Call OKButton_Click
    End If

End Sub

Private Sub txt2_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = 13 Then
        Call OKButton_Click
    End If

End Sub
