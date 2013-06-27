VERSION 5.00
Begin VB.Form frmSelectorSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selector Settings"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2550
   Icon            =   "frmSelectorSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   2550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraGeneral 
      Caption         =   "General:"
      Height          =   1875
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2355
      Begin VB.CheckBox chkAllowEdit 
         Caption         =   "Allow Edit"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1500
         Width           =   1215
      End
      Begin VB.CheckBox chkAutomaticClear 
         Caption         =   "Automatic Clear"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chkAutoFlash 
         Caption         =   "Auto Flash"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   540
         Width           =   1335
      End
      Begin VB.CheckBox chkAutoSelect 
         Caption         =   "Auto Select"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   1215
      End
      Begin VB.CheckBox chkAutoZoom 
         Caption         =   "Auto Zoom"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1140
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   3060
      Width           =   855
   End
   Begin VB.Frame FraSelectionSettings 
      Caption         =   "Selection Settings:"
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2445
      Begin VB.ComboBox ComBuffLevel 
         Height          =   315
         ItemData        =   "frmSelectorSettings.frx":6852
         Left            =   150
         List            =   "frmSelectorSettings.frx":687A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   1335
      End
      Begin VB.ComboBox txtSpatialOperation 
         Height          =   315
         ItemData        =   "frmSelectorSettings.frx":68BA
         Left            =   150
         List            =   "frmSelectorSettings.frx":68C1
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   540
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmSelectorSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event SettingsDone()

Private Sub cmdOK_Click()
    RaiseEvent SettingsDone
    Me.Hide
End Sub

Public Sub Init(bufflevel As Double, _
                sDE9IM As String, _
                bAutoZoom As Boolean, _
                bAutoSelect As Boolean, _
                bAutoFlash As Boolean, _
                bAutoClear As Boolean, _
                bEdit As Boolean)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        
        Dim keyArray() As Variant
        Dim element As Variant
        
        chkAutoFlash.Value = IIf(bAutoFlash, vbChecked, vbUnchecked)
        chkAutomaticClear.Value = IIf(bAutoClear, vbChecked, vbUnchecked)
        chkAutoSelect.Value = IIf(bAutoSelect, vbChecked, vbUnchecked)
        chkAutoZoom.Value = IIf(bAutoZoom, vbChecked, vbUnchecked)
        chkAllowEdit.Value = IIf(bEdit, vbChecked, vbUnchecked)
        
        FindIndexStrEx ComBuffLevel, CStr(bufflevel)
        '100     ComBuffLevel.ListIndex = 2
    
102     txtSpatialOperation.Clear
104     keyArray = DE9IM.Keys

106     For Each element In keyArray
108         txtSpatialOperation.AddItem element
        Next

        If Not sDE9IM = "" Then
            FindIndexStrEx txtSpatialOperation, sDE9IM
        Else

110         txtSpatialOperation.ListIndex = 5
        End If
    
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmSelectorSettings.init " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

