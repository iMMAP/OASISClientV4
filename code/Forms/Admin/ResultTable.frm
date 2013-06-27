VERSION 5.00
Begin VB.Form ResultTable 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prediction Table"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "ResultTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   7095
      Left            =   5160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   3240
      Picture         =   "ResultTable.frx":0C9E
      ToolTipText     =   "Make Another Prediction"
      Top             =   7440
      Width           =   4095
   End
   Begin VB.Image NewUp 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   4680
      Picture         =   "ResultTable.frx":7696
      Top             =   18000
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.Image NewDown 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   480
      Picture         =   "ResultTable.frx":E08E
      Top             =   18000
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2280
      Index           =   4
      Left            =   1800
      Top             =   4920
      Width           =   1530
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2280
      Index           =   3
      Left            =   1800
      Top             =   120
      Width           =   1530
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2280
      Index           =   2
      Left            =   120
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2280
      Index           =   1
      Left            =   1800
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2280
      Index           =   0
      Left            =   3480
      Top             =   2520
      Width           =   1530
   End
End
Attribute VB_Name = "ResultTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        Dim i As Integer

100     For i = 0 To 4
102         Image1(i).BorderStyle = 0
104     Next i

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.ResultTable.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Image2_Click()
'        '<EhHeader>
'        On Error GoTo Image2_Click_Err
'        '</EhHeader>
'100     RuneTable.Show (Form)
'102     RuneTable.Reset
'104     Unload Me
'        '<EhFooter>
'        Exit Sub
'
'Image2_Click_Err:
'        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.ResultTable.Image2_Click " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo Image2_MouseDown_Err
        '</EhHeader>
100 Image2.Picture = NewDown.Picture
        '<EhFooter>
        Exit Sub

Image2_MouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.ResultTable.Image2_MouseDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo Image2_MouseMove_Err
        '</EhHeader>
100 Image2.Picture = NewUp.Picture
        '<EhFooter>
        Exit Sub

Image2_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.ResultTable.Image2_MouseMove " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
