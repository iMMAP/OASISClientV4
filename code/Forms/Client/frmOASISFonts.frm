VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOASISFonts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Symbol Chooser"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4035
   Icon            =   "frmOASISFonts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   2910
      TabIndex        =   7
      Top             =   3000
      Width           =   1080
   End
   Begin VB.Frame FraAvailableFonts 
      Caption         =   "Available Fonts"
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4035
      Begin VB.ComboBox ComFonts 
         Height          =   315
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   315
         Width           =   3840
      End
      Begin VB.ComboBox comFontSize 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3150
         TabIndex        =   3
         Text            =   "12"
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ListView lvFonts 
         Height          =   2130
         Left            =   90
         TabIndex        =   2
         Top             =   690
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   3757
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Character"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   405
      Left            =   1800
      TabIndex        =   0
      Top             =   3000
      Width           =   1080
   End
   Begin VB.Label lblCurrentActive 
      Caption         =   "Current Active Symbol"
      Height          =   435
      Left            =   0
      TabIndex        =   6
      Top             =   2940
      Width           =   1035
   End
   Begin VB.Label lblLabel1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1080
      TabIndex        =   5
      Top             =   2910
      Width           =   645
   End
End
Attribute VB_Name = "frmOASISFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Property Get SymFontCharacter() As String
    SymFontCharacter = lblLabel1.caption
End Property

Public Property Get SymFontName() As String
    SymFontName = lblLabel1.Font.Name
End Property

Public Sub Init(SymFontName As String, SymFontCharacter As String)
    'Load the fonts into list1
    Dim x As Long
    Dim iCurrFont As Integer
        
    For x = 1 To Screen.FontCount
        If SymFontName = Screen.Fonts(x) Then iCurrFont = x - 1
        ComFonts.AddItem Screen.Fonts(x)
    Next

'    If ComFonts.ListCount > 0 Then ComFonts.ListIndex = iCurrFont

    For x = 0 To 255
        lvFonts.ListItems.Add Text:=Chr(x)
    Next
    
    lvFonts.ListItems.Item(Asc(SymFontCharacter) + 1).Selected = True
        
    lblLabel1.Font.Name = SymFontName
    lblLabel1.caption = SymFontCharacter
    ItemInBox SymFontName, ComFonts
    
End Sub

Private Sub cmdApply_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = "KILL"
    Unload Me
End Sub

Private Sub ComFonts_Click()

    If ComFonts.List(ComFonts.ListIndex) <> "" Then
        lvFonts.Font.Name = ComFonts.List(ComFonts.ListIndex)
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    Select Case UnloadMode
    
        Case vbFormControlMenu 'UnloadMode 0
            Me.Tag = "Kill"
            lblLabel1.caption = ""
            'form is being unloaded via the Close
            'command from the System menu
            'or by hitting the X in the upper right hand corner
        Case vbFormCode 'UnloadMode 1

            'Unload Me has been issued from code
        Case vbAppWindows 'UnloadMode 2

            'Windows itself is closing
        Case vbAppTaskManager 'UnloadMode 3

            'the Task Manager is closing the app
        Case vbFormMDIForm 'UnloadMod 4

            'an MDI child form is closing because
            'its parent form is closing
        Case vbFormOwner ' UnloadMode 5
            ' The owner of the form is closing
    
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not Len(Me.Tag) > 1 Then
        Cancel = 1
        Me.Hide
        Me.Tag = "1"
    Else
        lblLabel1.caption = ""
    End If
End Sub

Private Sub lvFonts_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Selected Then
        lblLabel1.caption = lvFonts.SelectedItem.Text
        If Not ComFonts.List(ComFonts.ListIndex) = "" Then
            lblLabel1.Font.Name = ComFonts.List(ComFonts.ListIndex)
        End If
    End If
End Sub
