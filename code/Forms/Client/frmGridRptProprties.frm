VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGridRptProprties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Grid Report"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4230
   Icon            =   "frmGridRptProprties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   465
      Left            =   2940
      TabIndex        =   9
      Top             =   4560
      Width           =   1245
   End
   Begin VB.Frame FraMapSettings 
      Caption         =   "Map Settings:"
      Height          =   855
      Left            =   30
      TabIndex        =   5
      Top             =   330
      Width           =   4155
      Begin VB.TextBox txtMapTitle 
         Height          =   285
         Left            =   900
         TabIndex        =   8
         Top             =   510
         Width           =   3195
      End
      Begin VB.CheckBox chkIncludeMap 
         Caption         =   "Include Map"
         Height          =   225
         Left            =   90
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblMapTitle 
         AutoSize        =   -1  'True
         Caption         =   "Map Title:"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   540
         Width           =   705
      End
   End
   Begin VB.CheckBox chkOnlyInclude 
      Caption         =   "Only include selected in grid"
      Height          =   285
      Left            =   3180
      TabIndex        =   3
      Top             =   1860
      Visible         =   0   'False
      Width           =   1005
   End
   Begin MSComctlLib.ListView lstFields 
      Height          =   3105
      Left            =   0
      TabIndex        =   2
      Top             =   1410
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   5477
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   390
      TabIndex        =   1
      Top             =   30
      Width           =   3795
   End
   Begin VB.Label lblIncludedFields 
      Caption         =   "Included Fields:"
      Height          =   195
      Left            =   30
      TabIndex        =   4
      Top             =   1230
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   345
   End
End
Attribute VB_Name = "frmGridRptProprties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bOkClicked As Boolean
Private mORs As ADODB.Recordset

Public Sub Init(RS As ADODB.Recordset, _
                Optional sGroups As String)
    Dim i As Integer
    bOkClicked = False
    Set mORs = RS
    
    For i = 0 To mORs.Fields.Count - 1
        
        lstFields.ListItems.Add , , mORs.Fields.Item(i).Name
            lstFields.ListItems.Item(lstFields.ListItems.Count).Checked = True
            
        If InStr(1, sGroups, mORs.Fields.Item(i).Name, vbTextCompare) > 0 Then
            lstFields.ListItems.Item(lstFields.ListItems.Count).Bold = True
        End If
        
        'lstFields.ListItems.Item(lstFields.ListItems.Count).
    Next
                                    
End Sub

Private Sub cmdOk_Click()
    bOkClicked = True
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
     
    If Not mORs Is Nothing Then

        For i = lstFields.ListItems.Count To 1 Step -1

            If Not lstFields.ListItems.Item(i).Checked Then
                mORs.Fields.Delete i - 1
            End If

        Next

        Set mORs = Nothing
        Cancel = 1
        Me.Visible = False
    End If
    
End Sub

Private Sub lstFields_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Bold Then
    Item.Checked = True
    End If
End Sub
