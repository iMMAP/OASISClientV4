VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmTextAnnoSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text Annotation Settings"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3525
   Icon            =   "frmTextAnnoSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox lstTexts 
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3240
      Width           =   3465
   End
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "Clear All"
      Height          =   315
      Left            =   30
      TabIndex        =   7
      Top             =   3600
      Width           =   1035
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   1260
      TabIndex        =   6
      Top             =   3600
      Width           =   1035
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   315
      Left            =   2460
      TabIndex        =   5
      Top             =   3600
      Width           =   1035
   End
   Begin VB.ListBox lstTexts1 
      Height          =   645
      Left            =   3240
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.TextBox txtAnnoText 
      Height          =   1875
      Left            =   30
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmTextAnnoSettings.frx":6852
      Top             =   1140
      Width           =   3465
   End
   Begin VB.Frame FraGeneralSettings 
      Caption         =   "General Settings"
      Height          =   915
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   3465
      Begin VB.Frame panColorBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   165
         Left            =   960
         TabIndex        =   13
         Top             =   450
         Width           =   495
      End
      Begin VB.Frame panColorFore 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000F&
         Height          =   165
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox ComFontSize 
         Height          =   315
         ItemData        =   "frmTextAnnoSettings.frx":6862
         Left            =   2700
         List            =   "frmTextAnnoSettings.frx":6890
         TabIndex        =   9
         Text            =   "12"
         Top             =   210
         Width           =   705
      End
      Begin VB.TextBox txtRotation 
         Height          =   285
         Left            =   2700
         TabIndex        =   4
         Text            =   "0"
         Top             =   540
         Width           =   420
      End
      Begin VB.CheckBox chkMultipleText 
         Caption         =   "Multiple Annotations"
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   660
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin MSComCtl2.UpDown udRotationAngle 
         Height          =   285
         Left            =   3120
         TabIndex        =   11
         Top             =   540
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtRotation"
         BuddyDispid     =   196619
         OrigLeft        =   1680
         OrigRight       =   1920
         OrigBottom      =   255
         Increment       =   5
         Max             =   180
         Min             =   -180
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblBackcolor 
         Caption         =   "Backcolor:"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblForecolor 
         Caption         =   "Forecolor:"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   210
         Width           =   705
      End
      Begin VB.Label lblAngle 
         AutoSize        =   -1  'True
         Caption         =   "Rotation:"
         Height          =   195
         Left            =   2010
         TabIndex        =   10
         Top             =   570
         Width           =   645
      End
      Begin VB.Label lblFontSize 
         Caption         =   "Font Size:"
         Height          =   225
         Left            =   1980
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label lblActiveAnnotation 
      AutoSize        =   -1  'True
      Caption         =   "Active Annotation:"
      Height          =   195
      Left            =   60
      TabIndex        =   18
      Top             =   3030
      Width           =   1305
   End
   Begin VB.Label lblAnnotationText 
      AutoSize        =   -1  'True
      Caption         =   "Annotation Text:"
      Height          =   195
      Left            =   60
      TabIndex        =   16
      Top             =   930
      Width           =   1170
   End
End
Attribute VB_Name = "frmTextAnnoSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event UpdateAnnoText()
Public Event GetAnnoTextProp()
Public Event DeleteAnnoText()
Public Event WinClosed()
Public Event ResetAll()

Private Sub cmdDelete_Click()
    RaiseEvent DeleteAnnoText
End Sub

Private Sub cmdRemoveAll_Click()
    RaiseEvent ResetAll
End Sub

Private Sub cmdUpdate_Click()
    RaiseEvent UpdateAnnoText
End Sub

Private Sub lstTexts_Click()
    RaiseEvent GetAnnoTextProp
End Sub

Private Sub panColorBack_Click()
    Dim c As New cCommonDialog
    
    c.ShowColor
    panColorBack.BackColor = c.Color

End Sub

Private Sub panColorFore_Click()
    Dim c As New cCommonDialog
    
    c.ShowColor
    panColorFore.BackColor = c.Color
        
End Sub
