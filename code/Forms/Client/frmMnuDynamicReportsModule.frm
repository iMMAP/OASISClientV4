VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmMnuDynamicReportsModule 
   BackColor       =   &H0050C0A4&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   FillColor       =   &H0050C0A4&
   ForeColor       =   &H0050C0A4&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7665
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3540
      _cx             =   6244
      _cy             =   13520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
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
      GridRows        =   8
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   5292196
      FrameShadow     =   5292196
      FloodStyle      =   1
      _GridInfo       =   $"frmMnuDynamicReportsModule.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.ListBox listGroup 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   1620
         Left            =   90
         TabIndex        =   8
         Top             =   1800
         Width           =   3360
      End
      Begin VB.ComboBox ComFilter 
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   7290
         Width           =   3360
      End
      Begin VB.ListBox Combo1 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   1035
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   375
         Width           =   3360
      End
      Begin VB.ListBox listQueries 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000005&
         Height          =   3150
         Left            =   90
         TabIndex        =   3
         Top             =   3735
         Width           =   3360
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Report Group"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   1515
         Width           =   3360
      End
      Begin VB.Label lblModuleFilter 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Report Filter"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   7005
         Width           =   3360
      End
      Begin VB.Label lblSelectData 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   3450
         Width           =   3360
      End
      Begin VB.Label lblSelectYour 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Module"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   3360
      End
   End
End
Attribute VB_Name = "frmMnuDynamicReportsModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event DatabaseClicked()
Public Event DRListClicked()
Public Event DRFilterClicked()
Public Event DRGroupClicked()

Private Sub Combo1_Click()
    RaiseEvent DatabaseClicked
End Sub

Private Sub ComFilter_Click()
    If Not ComFilter.Tag = "no event" Then
        RaiseEvent DRFilterClicked
    End If
End Sub

Private Sub listGroup_Click()
    RaiseEvent DRGroupClicked
End Sub

Private Sub listQueries_Click()
    RaiseEvent DRListClicked
End Sub
