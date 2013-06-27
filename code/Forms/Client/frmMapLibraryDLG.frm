VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMapLibraryDLG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Map Library"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraMapInformation 
      Caption         =   "Map Information:"
      Height          =   5835
      Left            =   4560
      TabIndex        =   27
      Top             =   60
      Width           =   4575
      Begin VB.Frame FraMapPreview 
         Caption         =   "Map Preview:"
         Height          =   2355
         Left            =   180
         TabIndex        =   29
         Top             =   240
         Width           =   4215
         Begin VB.PictureBox picMapPreview 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1995
            Left            =   120
            ScaleHeight     =   1965
            ScaleWidth      =   3945
            TabIndex        =   30
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.TextBox txtMapSummary 
         Alignment       =   2  'Center
         Height          =   3015
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2700
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2475
      TabIndex        =   12
      Top             =   5940
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   360
      Left            =   3510
      TabIndex        =   11
      Top             =   5940
      Width           =   990
   End
   Begin VB.Frame FraFrmMainDetails 
      Caption         =   "Map Details"
      Height          =   5820
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4425
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1020
         TabIndex        =   26
         Top             =   960
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy"
         Format          =   80805891
         CurrentDate     =   41060
      End
      Begin VB.TextBox txtMapInfo 
         Height          =   285
         Index           =   5
         Left            =   990
         TabIndex        =   20
         Top             =   2460
         Width           =   3210
      End
      Begin VB.TextBox txtMapInfo 
         Height          =   285
         Index           =   4
         Left            =   990
         TabIndex        =   18
         Top             =   2100
         Width           =   3210
      End
      Begin VB.TextBox txtMapInfo 
         Height          =   1230
         Index           =   6
         Left            =   990
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2865
         Width           =   3210
      End
      Begin VB.TextBox txtMapInfo 
         Height          =   285
         Index           =   3
         Left            =   990
         TabIndex        =   8
         Top             =   1695
         Width           =   3210
      End
      Begin VB.TextBox txtMapInfo 
         Height          =   285
         Index           =   2
         Left            =   990
         TabIndex        =   6
         Top             =   1335
         Width           =   3210
      End
      Begin VB.TextBox txtMapInfo 
         Height          =   285
         Index           =   1
         Left            =   990
         TabIndex        =   4
         Top             =   630
         Width           =   3210
      End
      Begin VB.TextBox txtMapInfo 
         Height          =   285
         Index           =   0
         Left            =   990
         TabIndex        =   2
         Top             =   270
         Width           =   3210
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Created:"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblGISDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EPSG:"
         Height          =   195
         Index           =   7
         Left            =   2340
         TabIndex        =   24
         Top             =   5340
         Width           =   435
      End
      Begin VB.Label lblGISDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scale:"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   23
         Top             =   5340
         Width           =   435
      End
      Begin VB.Label lblGISDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Center Y:"
         Height          =   195
         Index           =   5
         Left            =   2340
         TabIndex        =   22
         Top             =   4980
         Width           =   690
      End
      Begin VB.Label lblGISDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Center X:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   21
         Top             =   4980
         Width           =   690
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact:"
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   19
         Top             =   2460
         Width           =   630
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   17
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label lblGISDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y Max:"
         Height          =   195
         Index           =   3
         Left            =   2340
         TabIndex        =   16
         Top             =   4620
         Width           =   495
      End
      Begin VB.Label lblGISDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y Min:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Top             =   4620
         Width           =   435
      End
      Begin VB.Label lblGISDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X Max:"
         Height          =   195
         Index           =   1
         Left            =   2340
         TabIndex        =   14
         Top             =   4260
         Width           =   495
      End
      Begin VB.Label lblGISDetails 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X Min:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   4260
         Width           =   435
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source:"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   9
         Top             =   2100
         Width           =   555
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "URL:"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   7
         Top             =   1740
         Width           =   345
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   5
         Top             =   1380
         Width           =   765
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Author / Created By:"
         Height          =   420
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   930
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Map Name:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   315
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmMapLibraryDLG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event SaveClicked()
Public Event CancelClicked()

Public bOK As Boolean
Public bPreview As Boolean

Private Sub cmdCancel_Click()
    bOK = False
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    
    If Not bPreview Then
    If Len(txtMapInfo(0).Text) > 0 And Len(txtMapInfo(6).Text) > 0 Then
        If MsgBox("Do you want to save the map with these settings?", vbYesNo, "OASIS Map Library") = vbNo Then Exit Sub
    Else
        MsgBox "You have to fill in Name & Description of the map before saving."
        Exit Sub
    End If
    
    bOK = True
    End If
    
    Me.Hide
End Sub

Public Sub populatePreview()
    txtMapSummary.Text = ""
    txtMapSummary.Text = txtMapSummary.Text & "Map Name:" & " " & txtMapInfo(0).Text & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & "Created by:" & " " & txtMapInfo(1).Text & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & "Date Created:" & " " & DTPicker1.Value & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & "Copyright:" & " " & txtMapInfo(2).Text & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & "URL:" & " " & txtMapInfo(3).Text & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & "Source:" & " " & txtMapInfo(4).Text & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & "Contact:" & " " & txtMapInfo(5).Text & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & "Description:" & " " & txtMapInfo(6).Text & vbCrLf

    txtMapSummary.Text = txtMapSummary.Text & "-----------GEOGRAPHIC DETAILS--------" & vbCrLf

    txtMapSummary.Text = txtMapSummary.Text & lblGISDetails(0).caption & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & lblGISDetails(1).caption & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & lblGISDetails(2).caption & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & lblGISDetails(3).caption & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & lblGISDetails(4).caption & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & lblGISDetails(5).caption & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & lblGISDetails(6).caption & vbCrLf
    txtMapSummary.Text = txtMapSummary.Text & lblGISDetails(7).caption

End Sub

Public Sub Init(oGIS As XGIS_Viewer)

End Sub

Private Sub Form_Load()
    DebugPrint "Loading...."
    
End Sub
