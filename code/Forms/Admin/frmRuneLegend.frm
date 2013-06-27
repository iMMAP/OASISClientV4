VERSION 5.00
Begin VB.Form Legend 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Legend"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "frmRuneLegend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmRuneLegend.frx":0C9E
      Top             =   2520
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblDevelopedBy 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Developed by IMMAP INC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1245
      Left            =   1800
      TabIndex        =   2
      Top             =   5700
      Width           =   3225
   End
   Begin VB.Image Rt 
      Height          =   390
      Left            =   2160
      Picture         =   "frmRuneLegend.frx":1039
      Top             =   960
      Width           =   210
   End
   Begin VB.Image Lt 
      Height          =   390
      Left            =   120
      Picture         =   "frmRuneLegend.frx":14F3
      Top             =   960
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   480
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "Legend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CountC As Integer
Dim Def(24) As String
Dim RuneLegend As String
Dim Com As String

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     Image1.Picture = RuneTable.Cards(25).Picture

102     CountC = -1

104     RuneLegend = ""
106     Com = "This symbol indicates "

108     Def(0) = "the mouth of Odin, the God Of Wisdom."
110     Def(1) = "the breast of Earth Goddess."
112     Def(2) = "the awakening sound of the dawn."
114     Def(3) = "ice, Unbearable and extreme winter."
116     Def(4) = "two hands raised to the Heaven, The Symbol of Godly lookafter."
118     Def(5) = "the horns of cattle."
120     Def(6) = "married couples, Reunion, Godly gift."
122     Def(7) = "very strong icestorm."
124     Def(8) = "pregnant Earth Goddess."
126     Def(9) = "the conifer tree where Odin was hanged upside down."
128     Def(10) = "the seasons."
130     Def(11) = "flame(torch), The symbol of Nord's love, the God of Love."
132     Def(12) = "a broad lake, Water flow, the bath of a young child."
134     Def(13) = "your image in a mirror."
136     Def(14) = "a pair of horses, A combination of two forces heading ttowards a same goal."
138     Def(15) = "two swords, War, Opponent, Foe, Breaking apart."
140     Def(16) = "rituals and rites carried along generation, Hereditary wealth."
142     Def(17) = "the gourd in which Rune beads are kept. The word 'PERTHRU' denotes the good deeds of your prevoius lives."
144     Def(18) = "Thor, the God of Thunder and Lightning, Cactus plant."
146     Def(19) = "journey in a chariot"
148     Def(20) = "solar energy that melts ice."
150     Def(21) = "Tyre, the God of Martial Arts, Courage, Brave acts."
152     Def(22) = "Wild bull."
154     Def(23) = "flowered trees and creeers in the begining of spring."
156     Def(24) = "the oddes of fate."

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.Legend.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Lt_Click()
        '<EhHeader>
        On Error GoTo Lt_Click_Err
        '</EhHeader>

100     If CountC > 0 Then
102         CountC = CountC - 1
104         Image1.Picture = RuneTable.Cards(CountC).Picture
106         Text1.Text = Com + Def(CountC)
        End If

        '<EhFooter>
        Exit Sub

Lt_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.Legend.Lt_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Rt_Click()
        '<EhHeader>
        On Error GoTo Rt_Click_Err
        '</EhHeader>

100     If CountC < 24 Then
102         CountC = CountC + 1
104         Image1.Picture = RuneTable.Cards(CountC).Picture
106         Text1.Text = Com + Def(CountC)
        End If

        '<EhFooter>
        Exit Sub

Rt_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.Legend.Rt_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

