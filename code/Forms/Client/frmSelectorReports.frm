VERSION 5.00
Begin VB.Form frmSelectorReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selector Reports"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3780
   Icon            =   "frmSelectorReports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkIncludeMap 
      Caption         =   "Include Map"
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   360
      Width           =   1245
   End
   Begin VB.TextBox txtMapTitle 
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Top             =   660
      Width           =   2685
   End
   Begin VB.CheckBox chkIncludeCentroid 
      Caption         =   "Include Centroid"
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   1260
      Width           =   1485
   End
   Begin VB.CheckBox chkIncludeLength 
      Caption         =   "Include Length"
      Height          =   225
      Left            =   1680
      TabIndex        =   4
      Top             =   1020
      Width           =   1605
   End
   Begin VB.CheckBox chkIncludeArea 
      Caption         =   "Include Area"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   1260
      Width           =   1305
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   60
      Width           =   2655
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   285
      Left            =   2580
      TabIndex        =   1
      Top             =   1620
      Width           =   1005
   End
   Begin VB.CheckBox chkIncludeGeo 
      Caption         =   "Include Geo ID"
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   990
      Width           =   1575
   End
   Begin VB.Label lblMapTitle 
      AutoSize        =   -1  'True
      Caption         =   "Map Title:"
      Height          =   195
      Left            =   30
      TabIndex        =   9
      Top             =   690
      Width           =   705
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Report Title:"
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmSelectorReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event DoPrint()

Private Sub cmdPrint_Click()
    RaiseEvent DoPrint
    Me.Hide
End Sub
