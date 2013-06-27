VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmMnuDynamicDataModule 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2655
      _cx             =   4683
      _cy             =   7435
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   12632256
      ForeColor       =   5292196
      FloodColor      =   6553600
      ForeColorDisabled=   5292196
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   5
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   255
      FrameShadow     =   5292196
      FloodStyle      =   1
      _GridInfo       =   $"frmMnuDynamicDataModule.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.ListBox listDatabases 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   1395
         Left            =   90
         TabIndex        =   4
         Top             =   360
         Width           =   2475
      End
      Begin VB.ListBox lstDataElements 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   1980
         Left            =   90
         TabIndex        =   3
         Top             =   2100
         Width           =   2475
      End
      Begin VB.Label lblSelectYour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   2475
      End
      Begin VB.Label lblSelectData 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " Module Component"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   90
         TabIndex        =   1
         Top             =   1830
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frmMnuDynamicDataModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event DDDatabaseClicked()
Public Event DDTableClicked()

Private Sub ListDatabases_Click()
    RaiseEvent DDDatabaseClicked
End Sub

Private Sub lstDataElements_Click()
    RaiseEvent DDTableClicked
End Sub

