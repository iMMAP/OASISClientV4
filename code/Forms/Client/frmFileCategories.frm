VERSION 5.00
Begin VB.Form frmFileCategories 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "OASIS Attachments"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   1560
      MaxLength       =   128
      TabIndex        =   3
      Top             =   60
      Width           =   3015
   End
   Begin VB.ComboBox ComCategory 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Choose category:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label lblAddA 
      AutoSize        =   -1  'True
      Caption         =   "Add a title to you file:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "frmFileCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    g_sFileCategory = ComCategory.List(ComCategory.ListIndex)
    g_sFileTitle = txtTitle.Text
    Unload Me
End Sub

Private Sub ComCategory_Click()
    g_sFileCategory = ComCategory.List(ComCategory.ListIndex)
End Sub

Private Sub Form_Load()

    g_sFileTitle = txtTitle.Text
    
    With ComCategory
        .Clear
        .AddItem "Document"
        .AddItem "Images"
        .AddItem "Multimedia"
        .AddItem "Medical"
        .AddItem "Maps"
        .AddItem "Data"
        .AddItem "Other"
        .ListIndex = 0
    End With

    

End Sub
