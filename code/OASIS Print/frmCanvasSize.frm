VERSION 5.00
Begin VB.Form frmCanvasSize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Size"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4020
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4020
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton OpType 
      Caption         =   "Millimeters"
      Height          =   240
      Index           =   2
      Left            =   2205
      TabIndex        =   7
      Top             =   765
      Width           =   1500
   End
   Begin VB.OptionButton OpType 
      Caption         =   "Inches"
      Height          =   240
      Index           =   1
      Left            =   2205
      TabIndex        =   6
      Top             =   495
      Width           =   1500
   End
   Begin VB.OptionButton OpType 
      Caption         =   "Pixels"
      Height          =   240
      Index           =   0
      Left            =   2205
      TabIndex        =   5
      Top             =   225
      Value           =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton BtnOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1260
      TabIndex        =   4
      Top             =   1125
      Width           =   1590
   End
   Begin VB.TextBox TxtWidth 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1035
      TabIndex        =   1
      Text            =   "640"
      Top             =   630
      Width           =   690
   End
   Begin VB.TextBox TxtHeight 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1035
      TabIndex        =   0
      Text            =   "480"
      Top             =   225
      Width           =   690
   End
   Begin VB.Label Label2 
      Caption         =   "Width:"
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Top             =   675
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Height:"
      Height          =   240
      Left            =   135
      TabIndex        =   2
      Top             =   270
      Width           =   780
   End
End
Attribute VB_Name = "frmCanvasSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim myH As Single
Dim myW As Single

Private Sub BtnOk_Click()
    On Error GoTo Err1

    frmMainPrint.OASISDrawObj1.CanvasHeight = myH
    frmMainPrint.OASISDrawObj1.CanvasWidth = myW
    Unload Me
    Exit Sub
Err1:

    frmMainPrint.OASISDrawObj1.CanvasHeight = 480
    frmMainPrint.OASISDrawObj1.CanvasWidth = 640
    Unload Me
End Sub

Private Sub Form_Load()
    myH = frmMainPrint.OASISDrawObj1.CanvasHeight
    myW = frmMainPrint.OASISDrawObj1.CanvasWidth
    TxtHeight.Text = myH
    TxtWidth.Text = myW
End Sub

Private Sub OpType_Click(Index As Integer)

    Select Case Index

        Case 0
            TxtWidth.Text = myW
            TxtHeight.Text = myH

        Case 1
            TxtWidth.Text = ScaleX(myW, vbPixels, vbInches)
            TxtHeight.Text = ScaleY(myH, vbPixels, vbInches)

        Case 2
            TxtWidth.Text = ScaleX(myW, vbPixels, vbMillimeters)
            TxtHeight.Text = ScaleY(myH, vbPixels, vbMillimeters)
    End Select

End Sub

Private Sub TxtHeight_Change()
    On Error Resume Next

    If OpType(1).value = True Then
        myH = ScaleY(TxtHeight.Text, vbInches, vbPixels)
    ElseIf OpType(2).value = True Then
        myH = ScaleY(TxtHeight.Text, vbMillimeters, vbPixels)
    Else
        myH = TxtHeight.Text
    End If

End Sub

Private Sub TxtWidth_Change()
    On Error Resume Next

    If OpType(1).value = True Then
        myW = ScaleX(TxtWidth.Text, vbInches, vbPixels)
    ElseIf OpType(2).value = True Then
        myW = ScaleX(TxtWidth.Text, vbMillimeters, vbPixels)
    Else
        myW = TxtWidth.Text
    End If

End Sub

