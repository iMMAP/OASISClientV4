VERSION 5.00
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Begin VB.Form frmPrintTextProperties 
   BackColor       =   &H0050C0A4&
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   2880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraPreview 
      BackColor       =   &H0050C0A4&
      Caption         =   "Preview"
      Height          =   1515
      Left            =   30
      TabIndex        =   14
      Top             =   1740
      Width           =   2805
      Begin VB.Label lblJoeDonahue 
         Alignment       =   2  'Center
         BackColor       =   &H0050C0A4&
         Caption         =   "Joe Donahue Red Fox"
         Height          =   1215
         Left            =   90
         TabIndex        =   15
         Top             =   210
         Width           =   2625
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0050C0A4&
      Caption         =   "Frame1"
      Height          =   1665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2865
      Begin VB.CommandButton cmdFonts 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   1740
         TabIndex        =   9
         Top             =   1170
         Width           =   465
      End
      Begin VB.TextBox txtTextTop 
         Height          =   285
         Index           =   0
         Left            =   420
         TabIndex        =   8
         Text            =   "0.75"
         Top             =   495
         Width           =   465
      End
      Begin VB.TextBox txtTextLeft 
         Height          =   285
         Index           =   0
         Left            =   420
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   195
         Width           =   465
      End
      Begin VB.TextBox txtTextRight 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Text            =   "-0.75"
         Top             =   825
         Width           =   465
      End
      Begin VB.TextBox txtTextBottom 
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   5
         Text            =   "3.00"
         Top             =   1155
         Width           =   465
      End
      Begin VB.CheckBox chkUnderline 
         BackColor       =   &H0050C0A4&
         Caption         =   "Underline"
         Height          =   195
         Index           =   0
         Left            =   1110
         TabIndex        =   3
         Top             =   240
         Width           =   1005
      End
      Begin VB.CheckBox chkStrikeThrough 
         BackColor       =   &H0050C0A4&
         Caption         =   "Strike Through"
         Height          =   225
         Index           =   0
         Left            =   1110
         TabIndex        =   2
         Top             =   480
         Width           =   1365
      End
      Begin VB.ComboBox ComJustify 
         Height          =   315
         Index           =   0
         ItemData        =   "frmPrintTextProperties.frx":0000
         Left            =   1110
         List            =   "frmPrintTextProperties.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   1545
      End
      Begin XpressEditorsLibCtl.dxColorEdit dxColorTXT 
         Height          =   315
         Index           =   0
         Left            =   1110
         OleObjectBlob   =   "frmPrintTextProperties.frx":003D
         TabIndex        =   4
         Top             =   1110
         Width           =   525
      End
      Begin VB.Label lblMap 
         AutoSize        =   -1  'True
         BackColor       =   &H0050C0A4&
         Caption         =   "Bottom"
         Height          =   195
         Index           =   23
         Left            =   0
         TabIndex        =   13
         Top             =   1215
         Width           =   495
      End
      Begin VB.Label lblMap 
         AutoSize        =   -1  'True
         BackColor       =   &H0050C0A4&
         Caption         =   "Right"
         Height          =   195
         Index           =   24
         Left            =   60
         TabIndex        =   12
         Top             =   885
         Width           =   375
      End
      Begin VB.Label lblMap 
         AutoSize        =   -1  'True
         BackColor       =   &H0050C0A4&
         Caption         =   "Top"
         Height          =   195
         Index           =   25
         Left            =   90
         TabIndex        =   11
         Top             =   555
         Width           =   285
      End
      Begin VB.Label lblMap 
         AutoSize        =   -1  'True
         BackColor       =   &H0050C0A4&
         Caption         =   "Left"
         Height          =   195
         Index           =   26
         Left            =   90
         TabIndex        =   10
         Top             =   255
         Width           =   270
      End
   End
End
Attribute VB_Name = "frmPrintTextProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

