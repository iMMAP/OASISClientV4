VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmAccess 
   Caption         =   "Security And Access"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   15690
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   5835
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15690
      _cx             =   27675
      _cy             =   10292
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   2
      ChildSpacing    =   1
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
      GridRows        =   2
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmAccess.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   540
         Left            =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   5265
         Width           =   15630
         _cx             =   27570
         _cy             =   953
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
         Begin VB.CommandButton Command6 
            Caption         =   "Export"
            Height          =   465
            Left            =   11595
            TabIndex        =   8
            Top             =   45
            Width           =   1770
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Synchronise"
            Height          =   465
            Left            =   13590
            TabIndex        =   7
            Top             =   45
            Width           =   1770
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Access Per District"
            Height          =   465
            Left            =   3600
            TabIndex        =   6
            Top             =   45
            Width           =   1770
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Access Per Gov"
            Height          =   465
            Left            =   5595
            TabIndex        =   5
            Top             =   45
            Width           =   1770
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Check Position"
            Height          =   465
            Left            =   7590
            TabIndex        =   4
            Top             =   45
            Width           =   1770
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Add Position from Map"
            Height          =   465
            Left            =   9600
            TabIndex        =   3
            Top             =   45
            Width           =   1770
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   5220
         Left            =   30
         OleObjectBlob   =   "frmAccess.frx":0045
         TabIndex        =   1
         Top             =   30
         Width           =   15630
      End
   End
End
Attribute VB_Name = "frmAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If

End Sub
