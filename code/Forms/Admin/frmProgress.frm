VERSION 5.00
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Begin VB.Form frmProgress 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Task Progress"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   405
   End
   Begin CONTROLSLibCtl.dxLabel DxLPosition 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4395
      _Version        =   0
      _cx             =   7752
      _cy             =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Processing OASIS Server Request..."
      BackStyle       =   1
      BackColor       =   -2147483633
      ForeColor       =   0
      LabelStyle      =   0
      Label3dStyle    =   2
      Label3dOrientation=   4
      Label3dDepth    =   0
      PenWidth        =   1
      Angle           =   0
      ShadowColor     =   8421504
   End
   Begin CONTROLSLibCtl.dxProgressBar dxProgressBar 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _Version        =   65536
      _cx             =   7858
      _cy             =   617
      ForeColor       =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MinPos          =   0
      MaxPos          =   100
      Pos             =   50
      Step            =   10
      ShowText        =   0   'False
      Orientation     =   0
      StartColor      =   128
      EndColor        =   16777215
      DrawBorderStyle =   1
      ShowTextStyle   =   0
      DrawBarStyle    =   3
      DrawBarBorderStyle=   0
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
        '<EhHeader>
        On Error GoTo Form_DblClick_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

Form_DblClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmProgress.Form_DblClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

