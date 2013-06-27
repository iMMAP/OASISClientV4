VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintLoader 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   1680
      Top             =   1380
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3200
      Left            =   0
      ScaleHeight     =   3165
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtLoading 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   2235
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   4335
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   60
         TabIndex        =   1
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblProgress 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading OASIS Print Utilities..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmPrintLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ProgressBar1.value = 0
    Timer1.Enabled = True
End Sub

Private Sub Form_LostFocus()
    Me.SetFocus
End Sub

Private Sub Timer1_Timer()
    If ProgressBar1.value = 100 Then
        ProgressBar1.value = 0
    Else
        If ProgressBar1.value > 90 Then ProgressBar1.value = 92.5
        ProgressBar1.value = ProgressBar1.value + 7.5
    End If
End Sub
