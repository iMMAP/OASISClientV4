VERSION 5.00
Begin VB.Form frmMapPreviewOptions 
   Caption         =   "Options"
   ClientHeight    =   4785
   ClientLeft      =   1320
   ClientTop       =   1335
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4785
   ScaleWidth      =   6690
   Begin VB.TextBox txtPageHeight 
      Height          =   360
      Left            =   1845
      TabIndex        =   11
      Top             =   4260
      Width           =   1035
   End
   Begin VB.TextBox txtPageWidth 
      Height          =   330
      Left            =   1845
      TabIndex        =   9
      Top             =   3765
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Caption         =   "Orientation"
      Height          =   1245
      Left            =   3795
      TabIndex        =   5
      Top             =   2280
      Width           =   1755
      Begin VB.OptionButton optOrientation 
         Caption         =   "Landscape"
         Height          =   315
         Index           =   1
         Left            =   270
         TabIndex        =   7
         Top             =   765
         Width           =   1155
      End
      Begin VB.OptionButton optOrientation 
         Caption         =   "Portraint"
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   6
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.CheckBox chkGrid 
      Caption         =   "Show Grid"
      Height          =   375
      Left            =   3855
      TabIndex        =   4
      Top             =   525
      Value           =   1  'Checked
      Width           =   1110
   End
   Begin VB.ListBox lstPageSize 
      Height          =   2985
      Left            =   270
      TabIndex        =   3
      Top             =   600
      Width           =   2580
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   5250
      TabIndex        =   1
      Top             =   4230
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   405
      Left            =   3780
      TabIndex        =   0
      Top             =   4245
      Width           =   1230
   End
   Begin VB.Label lblPageHeight 
      AutoSize        =   -1  'True
      Caption         =   "PageHeight (Twips)"
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   4350
      Width           =   1395
   End
   Begin VB.Label lblPageWidth 
      AutoSize        =   -1  'True
      Caption         =   "PageWidth (Twips)"
      Height          =   195
      Left            =   255
      TabIndex        =   8
      Top             =   3825
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Page Size"
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   315
      Width           =   720
   End
End
Attribute VB_Name = "frmMapPreviewOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
End Sub


Private Sub cmdOK_Click()
    On Error Resume Next
Dim vp As Control
Set vp = frmDesigner.ViewPro1

    'Show grid
    If chkGrid.Value = 1 Then
        frmDesigner.ShowGrid = 1
    Else
        frmDesigner.ShowGrid = 0
    End If
    
    'Orientation
    If optOrientation(0).Value = True Then
        vp.Orientation = 0
    End If
    If optOrientation(1).Value = True Then
        vp.Orientation = 1
    End If
    
    'Page size
    vp.PageSize = lstPageSize.ListIndex + 1
    If lstPageSize.ListIndex = 13 Then
        vp.PageWidth = Val(txtPageWidth.Text)
        vp.PageHeight = Val(txtPageHeight.Text)
    End If
    
    
    
    vp.UpdateDoc
    
    Unload Me
    
    
End Sub


Private Sub Form_Load()
    On Error Resume Next
Dim vp As Control

Set vp = frmDesigner.ViewPro1

    'Orientation
    If vp.Orientation = 0 Then
        optOrientation(0).Value = True
    Else
        optOrientation(1).Value = True
    End If
    
    'Show grid
    If frmDesigner.ShowGrid Then
        chkGrid.Value = 1
    Else
        chkGrid.Value = 0
    End If
    
    
    
    'Custom page size boxes
    txtPageWidth.Text = vp.PageWidth
    txtPageHeight.Text = vp.PageHeight
    If vp.PageSize <> 14 Then
        txtPageWidth.Visible = False
        txtPageHeight.Visible = False
        lblPageWidth.Visible = False
        lblPageHeight.Visible = False
    End If
    
    
    'Page size
    lstPageSize.AddItem "1 - Letter, 8 1/2 x 11 in."
    lstPageSize.AddItem "2 - Legal, 8 1/2 x 14 in."
    lstPageSize.AddItem "3 - Executive, 7 1/2 x 10 1/2 in."
    lstPageSize.AddItem "4 - Tabloid, 11 x 17 in."
    lstPageSize.AddItem "5 - Ledger, 17 x 11 in."
    lstPageSize.AddItem "6 - Statement, 5 1/2 x 8 1/2 in."
    lstPageSize.AddItem "7 - Folio, 8 1/2 x 13 in."
    lstPageSize.AddItem "8 - A3, 297 x 420 mm"
    lstPageSize.AddItem "9 - A4, 210 x 297 mm"
    lstPageSize.AddItem "10 - A5, 148 x 210 mm"
    lstPageSize.AddItem "11 - B4, 250 x 354 mm"
    lstPageSize.AddItem "12 - B5, 182 x 257 mm"
    lstPageSize.AddItem "13 - Quarto, 215 x 275 mm"
    lstPageSize.AddItem "14 - Custom Size"
    lstPageSize.ListIndex = vp.PageSize - 1


                
    
    
End Sub


Private Sub lstPageSize_Click()
    On Error Resume Next
    If lstPageSize.ListIndex = 13 Then
        txtPageWidth.Visible = True
        txtPageHeight.Visible = True
        lblPageWidth.Visible = True
        lblPageHeight.Visible = True
    Else
        txtPageWidth.Visible = False
        txtPageHeight.Visible = False
        lblPageWidth.Visible = False
        lblPageHeight.Visible = False
    End If
    
    
    
End Sub


