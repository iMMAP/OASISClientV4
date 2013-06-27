VERSION 5.00
Begin VB.Form frmViewer 
   Caption         =   "Environment Variable Viewer"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   3375
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "List All"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ListBox lstVariables 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdList_Click()
    Dim i As Integer
    Dim Buffer As String
    
    lstVariables.Clear
    Buffer = "a" ' To make sure that it enters the loop
    i = 1
    Do Until Buffer = "" ' If the Index specified in Environ(i) is non existant then it will return an emtpy string, we can use this to exit the loop
        Buffer = Environ(i) ' Set Buffer to Environment Variable(i), you can also use a string to find the value for a specific variable, for example BUFFER = ENVIRON("TMP")
        lstVariables.AddItem (Buffer) 'Display
        i = i + 1
    Loop
End Sub

Private Sub cmdSearch_Click() ' This will display the value for the specified Variable, or "" if it doesnt exist
    Dim Buffer As String
    
    lstVariables.Clear
    Buffer = Environ(txtSearch.Text)
    If Buffer = "" Then
        Buffer = "Variable non-existant or empty"
    End If
    lstVariables.AddItem (Buffer)
    txtSearch.Text = ""
End Sub

Private Sub Form_Resize()
    lstVariables.Left = 0
    lstVariables.Width = Me.Width
    lstVariables.Top = 0
End Sub
