VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Begin VB.Form frmMnuOASISProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Profile"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3705
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   3705
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   5565
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3705
      _cx             =   6535
      _cy             =   9816
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
      Align           =   5
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
      Begin XpressEditorsLibCtl.dxHyperLinkEdit dxHyperLinkEdit4 
         Height          =   780
         Left            =   405
         OleObjectBlob   =   "frmMnuOASISProfile.frx":0000
         TabIndex        =   4
         Top             =   2970
         Visible         =   0   'False
         Width           =   2805
      End
      Begin XpressEditorsLibCtl.dxHyperLinkEdit dxHyperLinkEdit3 
         Height          =   780
         Left            =   405
         OleObjectBlob   =   "frmMnuOASISProfile.frx":0A5B
         TabIndex        =   3
         Top             =   2070
         Visible         =   0   'False
         Width           =   2805
      End
      Begin XpressEditorsLibCtl.dxHyperLinkEdit dxHyperLinkEdit2 
         Height          =   780
         Left            =   405
         OleObjectBlob   =   "frmMnuOASISProfile.frx":14C8
         TabIndex        =   2
         Top             =   1215
         Width           =   2805
      End
      Begin XpressEditorsLibCtl.dxHyperLinkEdit dxHyperLinkEdit1 
         Height          =   780
         Left            =   405
         OleObjectBlob   =   "frmMnuOASISProfile.frx":1D9F
         TabIndex        =   1
         Top             =   360
         Width           =   2805
      End
   End
End
Attribute VB_Name = "frmMnuOASISProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OASISProfileClicked()
Public Event OASISIntranetClicked()
Public Event OASISSettings()
Public Event OASISSupportCentre()

Private Sub dxHyperLinkEdit1_Click()
    RaiseEvent OASISProfileClicked
End Sub

Private Sub dxHyperLinkEdit2_Click()
    RaiseEvent OASISIntranetClicked
End Sub

Private Sub dxHyperLinkEdit3_Click()
    RaiseEvent OASISSettings
End Sub

Private Sub dxHyperLinkEdit4_Click()
    RaiseEvent OASISSupportCentre
End Sub

Private Sub Form_Load()
If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
End Sub
