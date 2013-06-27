VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{9309EA27-780F-4C3D-84E8-79DEB1CECE69}#5.0#0"; "ThemeRangePicker.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   4905
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6885
      _cx             =   12144
      _cy             =   8652
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
      AutoSizeChildren=   8
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
      GridRows        =   2
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmStart.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin OASISThemeRange.OASISThemeRangePicker OASISThemeRangePicker1 
         Height          =   4470
         Left            =   90
         TabIndex        =   2
         Top             =   345
         Width           =   6705
         _extentx        =   11827
         _extenty        =   7885
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   6705
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    OASISThemeRangePicker1.setNumberOfIntervals 3
    OASISThemeRangePicker1.SetInterval 1, vbYellow, 255
    OASISThemeRangePicker1.SetInterval 2, vbRed, 400
    OASISThemeRangePicker1.SetInterval 3, vbGreen, 1000
    OASISThemeRangePicker1.SetThemeStartColor vbBlue
    OASISThemeRangePicker1.Render
    Debug.Print "number of intervals: " & OASISThemeRangePicker1.GetNumberOfIntervals
    Debug.Print "start color: " & OASISThemeRangePicker1.GetThemeStartColor
    Debug.Print "interval col 1: " & OASISThemeRangePicker1.GetIntervalEndColor(1)
    Debug.Print "interval col 2: " & OASISThemeRangePicker1.GetIntervalEndColor(2)
    Debug.Print "interval col 3: " & OASISThemeRangePicker1.GetIntervalEndColor(3)
    'GetIntervalRatioValue
    Debug.Print "ratio col 1: " & OASISThemeRangePicker1.GetIntervalRatioValue(0)
    Debug.Print "ratio col 2: " & OASISThemeRangePicker1.GetIntervalRatioValue(2)
    Debug.Print "ratio col 3: " & OASISThemeRangePicker1.GetIntervalRatioValue(3)
    
End Sub

