VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.UserControl SubMeny 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ScaleHeight     =   4110
   ScaleWidth      =   2400
   Begin C1SizerLibCtl.C1Elastic elBtnHolder 
      Height          =   4200
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2445
      _cx             =   4313
      _cy             =   7408
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
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
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
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CommandButton cmdAddIncident 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   0
         Picture         =   "SubMeny.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   2400
      End
      Begin VB.CommandButton cmdViewIncidents 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   0
         Picture         =   "SubMeny.ctx":3C96
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2925
         Width           =   2400
      End
      Begin VB.CommandButton cmdCommonOperations 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   0
         Picture         =   "SubMeny.ctx":792C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3510
         Width           =   2400
      End
      Begin VB.CommandButton cmdOpsWizard 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   0
         Picture         =   "SubMeny.ctx":B5C2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   585
         Width           =   2400
      End
      Begin VB.CommandButton cmdRadioRoom 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   0
         Picture         =   "SubMeny.ctx":F258
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2340
         Width           =   2400
      End
      Begin VB.CommandButton cmdMyActivities 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   0
         Picture         =   "SubMeny.ctx":12EEE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1755
         Width           =   2400
      End
      Begin VB.CommandButton cmdPersonnel 
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   0
         Picture         =   "SubMeny.ctx":16B84
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1170
         Width           =   2400
      End
   End
End
Attribute VB_Name = "SubMeny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event OASISMnuPressed(oBtn As OASISMenuButton)

Private Sub cmdAddIncident_Click()
    RaiseEvent OASISMnuPressed(OASISMenuButton.Incidentwizard)
End Sub

Private Sub cmdCommonOperations_Click()
    RaiseEvent OASISMnuPressed(OASISMenuButton.LocationAnalysis)
End Sub

Private Sub cmdMyActivities_Click()
    RaiseEvent OASISMnuPressed(OASISMenuButton.Locationwizard)
End Sub

Private Sub cmdOpsWizard_Click()
    RaiseEvent OASISMnuPressed(OASISMenuButton.Activitieswizard)
End Sub

Private Sub cmdPersonnel_Click()
    RaiseEvent OASISMnuPressed(OASISMenuButton.Personellwizard)
End Sub

Private Sub cmdRadioRoom_Click()
    RaiseEvent OASISMnuPressed(OASISMenuButton.Radioroom)
End Sub

Private Sub cmdViewIncidents_Click()
    RaiseEvent OASISMnuPressed(OASISMenuButton.IncidentAnalysis)
End Sub
