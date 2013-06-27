VERSION 5.00
Begin VB.Form frmFolders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folders"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFolders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   """Copy To Clipboard EX"""
      Height          =   495
      Left            =   6255
      TabIndex        =   4
      Top             =   3600
      Width           =   1470
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy To Clipboard"
      Height          =   495
      Left            =   4770
      TabIndex        =   3
      Top             =   3600
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Explore"
      Height          =   495
      Left            =   3285
      TabIndex        =   2
      Top             =   3600
      Width           =   1470
   End
   Begin VB.TextBox txtFolder 
      Height          =   3255
      Left            =   3285
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   315
      Width           =   4440
   End
   Begin VB.ListBox lstFolder 
      Height          =   3570
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2955
   End
   Begin VB.Label lblSpecialFolders 
      Caption         =   "Special Folders:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   90
      Width           =   2895
   End
   Begin VB.Label lblPath 
      Caption         =   "Path:"
      Height          =   240
      Left            =   3285
      TabIndex        =   5
      Top             =   90
      Width           =   1545
   End
End
Attribute VB_Name = "frmFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>

100     Shell "explorer.exe " & Chr$(34) & txtFolder.Text & Chr$(34), vbNormalFocus

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFolders.Command1_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Command2_Click()

    Unload Me
    'End

End Sub

Private Sub Command3_Click()
        '<EhHeader>
        On Error GoTo Command3_Click_Err
        '</EhHeader>

100     Clipboard.Clear
102     Clipboard.SetText Chr$(34) & txtFolder.Text & Chr$(34), vbCFText

        '<EhFooter>
        Exit Sub

Command3_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFolders.Command3_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Command4_Click()
        '<EhHeader>
        On Error GoTo Command4_Click_Err
        '</EhHeader>

100     Clipboard.Clear
102     Clipboard.SetText txtFolder.Text, vbCFText

        '<EhFooter>
        Exit Sub

Command4_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFolders.Command4_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     With lstFolder
102         .AddItem "CD Burning Cache"
104         .itemData(.NewIndex) = 59&
106         .AddItem "Common Admin Tools"
108         .itemData(.NewIndex) = 47&
110         .AddItem "Common Application Data"
112         .itemData(.NewIndex) = 35&
114         .AddItem "Common Desktop"
116         .itemData(.NewIndex) = 25&
118         .AddItem "Common Document Templates"
120         .itemData(.NewIndex) = 45&
122         .AddItem "Common Favorites"
124         .itemData(.NewIndex) = 31&
126         .AddItem "Common My Documents"
128         .itemData(.NewIndex) = 46&
130         .AddItem "Common My Pictures"
132         .itemData(.NewIndex) = 54&
134         .AddItem "Common Program Files"
136         .itemData(.NewIndex) = 43&
138         .AddItem "Common Start Menu"
140         .itemData(.NewIndex) = 22&
142         .AddItem "Common Start Menu Programs"
144         .itemData(.NewIndex) = 23&
146         .AddItem "Common Startup"
148         .itemData(.NewIndex) = 24&
150         .AddItem "Fonts"
152         .itemData(.NewIndex) = 20&
154         .AddItem "Program Files"
156         .itemData(.NewIndex) = 38&
158         .AddItem "System32 Folder"
160         .itemData(.NewIndex) = 41&
162         .AddItem "System Folder"
164         .itemData(.NewIndex) = 37&
166         .AddItem "Themes"
168         .itemData(.NewIndex) = 56&
170         .AddItem "User Admin Tools"
172         .itemData(.NewIndex) = 48&
174         .AddItem "User Application Data"
176         .itemData(.NewIndex) = 26&
178         .AddItem "User Cookies"
180         .itemData(.NewIndex) = 33&
182         .AddItem "User Desktop"
184         .itemData(.NewIndex) = 16&
186         .AddItem "User Document Templates"
188         .itemData(.NewIndex) = 21&
190         .AddItem "User Favorites"
192         .itemData(.NewIndex) = 6&
194         .AddItem "User History"
196         .itemData(.NewIndex) = 34&
198         .AddItem "User Local Application Data"
200         .itemData(.NewIndex) = 28&
202         .AddItem "User My Documents"
204         .itemData(.NewIndex) = 5&
206         .AddItem "User My Music"
208         .itemData(.NewIndex) = 13&
210         .AddItem "User My Pictures"
212         .itemData(.NewIndex) = 39&
214         .AddItem "User Net Hood"
216         .itemData(.NewIndex) = 19&
218         .AddItem "User Print Hood"
220         .itemData(.NewIndex) = 27&
222         .AddItem "User Profile Folder"
224         .itemData(.NewIndex) = 40&
226         .AddItem "User Recent Documents"
228         .itemData(.NewIndex) = 8&
230         .AddItem "User SendTo"
232         .itemData(.NewIndex) = 9&
234         .AddItem "User Start Menu"
236         .itemData(.NewIndex) = 11&
238         .AddItem "UserStartMenuPrograms"
240         .itemData(.NewIndex) = 2&
242         .AddItem "User Startup"
244         .itemData(.NewIndex) = 7&
246         .AddItem "UserTempInternetFiles"
248         .itemData(.NewIndex) = 32&
250         .AddItem "Windows Folder"
252         .itemData(.NewIndex) = 36&
        End With

254     If Not g_sLanguage = "" Then
256         If Not m_Cnn.State = adStateClosed Then
258             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFolders.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub lstFolder_Click()
        '<EhHeader>
        On Error GoTo lstFolder_Click_Err
        '</EhHeader>

        Dim l As Long

100     l = CLng(lstFolder.itemData(lstFolder.ListIndex))
102     txtFolder = SpecialFolderPath(l)

        '<EhFooter>
        Exit Sub

lstFolder_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFolders.lstFolder_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

