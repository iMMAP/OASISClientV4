VERSION 5.00
Begin VB.Form frmAttachmentViewer 
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin OASISClient.OASISAttachments OASISAttachments1 
      Height          =   5655
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9975
   End
End
Attribute VB_Name = "frmAttachmentViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim ocon As New adodb.Connection

    OASISAttachments1.Init "www.oasiswebservice.org/upl/view.php?format=xmllist", ocon
End Sub
