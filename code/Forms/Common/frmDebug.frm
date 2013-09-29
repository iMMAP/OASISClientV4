VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmDebug 
   Caption         =   "Debug Window"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12840
   ClipControls    =   0   'False
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   12840
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12840
      _cx             =   22648
      _cy             =   4022
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
      BackColor       =   12648384
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   0
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
      GridRows        =   4
      GridCols        =   5
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmDebug.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CommandButton cmdAutoScroll 
         Caption         =   "Disable AutoScroll"
         Height          =   255
         Left            =   15
         TabIndex        =   4
         Top             =   2025
         Width           =   4245
      End
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Copy to Clipboard"
         Height          =   255
         Left            =   4260
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2025
         Width           =   4320
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1980
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Double click to view detail"
         Top             =   0
         Width           =   12840
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Clear"
         Height          =   255
         Left            =   8580
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2025
         Width           =   4260
      End
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function DebugPrinter(DebugStr As Variant)
        
    On Error Resume Next
        
    If FileExists(g_sAppPath & "\data\user\Sessions\temp_.log") = 1 Then
        'Petri: This needs to be fixed
        'Error: Invalid use of property
        'oDebug.LogFile g_sAppPath & "\data\user\Sessions\temp_.log"
    End If
    
    'Actual limit = 32,767
    If List1.ListCount > 10000 Then
        List1.Clear
        ''List1.RemoveItem 0
    End If

    List1.AddItem "[" & Time & "] " & CStr(DebugStr)
    
   ' oDebug.Log DebugStr
    
    If cmdAutoScroll.caption = "Disable AutoScroll" Then
        List1.ListIndex = List1.ListCount - 1
    End If

End Function

Private Sub cmdAutoScroll_Click()
    If cmdAutoScroll.caption = "Disable AutoScroll" Then
        cmdAutoScroll.caption = "Enable AutoScroll"
    Else
        cmdAutoScroll.caption = "Disable AutoScroll"
    End If
End Sub

Private Sub cmdClear_Click()
    List1.Clear
End Sub

Private Sub cmdExport_Click()

    Dim l As Long
    Dim sOutput As String
    
    l = 0
    Do Until l = List1.ListCount
        List1.ListIndex = l
        sOutput = sOutput & List1.Text & vbCrLf
        l = l + 1
    Loop
    Clipboard.Clear
    Clipboard.SetText sOutput, vbCFText
    MsgBox "Copied to clipboard"
    
End Sub

Private Sub Form_Load()
    Me.top = Screen.Height * 0.75
    Me.Height = Screen.Height * 0.2
    Me.left = Screen.Width * 0.25
    Me.Width = Screen.Width * 0.75
End Sub

Private Sub List1_DblClick()
    MsgBox List1.List(List1.ListIndex), vbInformation, "Debug data"
End Sub
