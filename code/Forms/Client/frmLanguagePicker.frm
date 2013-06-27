VERSION 5.00
Begin VB.Form frmLanguagePicker 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "OASIS Languages"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChangeLanguage 
      Caption         =   "Change Language"
      Height          =   330
      Left            =   1125
      TabIndex        =   1
      Top             =   450
      Width           =   1500
   End
   Begin VB.ComboBox ComLanguage 
      Height          =   315
      Left            =   45
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   2625
   End
End
Attribute VB_Name = "frmLanguagePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sNewLangugage As String

Private Sub cmdChangeLanguage_Click()
    Me.Hide
End Sub

Private Sub ComLanguage_Click()
    sNewLangugage = ComLanguage.List(ComLanguage.ListIndex)
End Sub

Private Sub Form_Load()
    Dim rslang As New ADODB.Recordset
    Dim i As Integer
    Dim j As Integer

    On Error GoTo Hell
    
    rslang.Open "Select * FROM lang", m_Cnn, adOpenDynamic, adLockReadOnly
    
    ComLanguage.Clear
    
    For i = 6 To rslang.Fields.Count - 1

        If rslang.Fields.Item(i).Name = sNewLangugage Then j = i - 6
        ComLanguage.AddItem rslang.Fields.Item(i).Name
    Next
    
    ComLanguage.ListIndex = j
    
    DoEvents
    
    sNewLangugage = ""
    
    GoTo BeyondHell
    
Hell:
    MsgBox "Could Not Read OASIS Language Table."
    On Error Resume Next
    DebugPrint Err.Description
    rslang.Close
    Set rslang = Nothing
    Unload Me
    Exit Sub
BeyondHell:
    rslang.Close
    Set rslang = Nothing
End Sub
