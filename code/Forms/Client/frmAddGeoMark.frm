VERSION 5.00
Begin VB.Form FrmAddBookMark 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Bookmark"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4050
      TabIndex        =   7
      Top             =   165
      Width           =   705
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   4050
      TabIndex        =   6
      Top             =   525
      Width           =   720
   End
   Begin VB.TextBox txtDesxription 
      Height          =   765
      Left            =   930
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   675
      Width           =   3045
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   930
      TabIndex        =   4
      Top             =   30
      Width           =   3045
   End
   Begin VB.ComboBox ComCategory 
      Height          =   315
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   345
      Width           =   3045
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   3
      Top             =   675
      Width           =   840
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Category:"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   345
      Width           =   675
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Index           =   0
      Left            =   375
      TabIndex        =   0
      Top             =   75
      Width           =   465
   End
End
Attribute VB_Name = "FrmAddBookMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mConn As ADODB.Connection
Private m_x As Double
Private m_y As Double
Private M_z As Double
Private m_sMapName As String

Public Sub Init(oConn As ADODB.Connection, x As Double, y As Double, z As Double, sMapName As String)
    Set mConn = oConn
    m_x = x
    m_y = y
    M_z = z
    m_sMapName = sMapName
End Sub

Private Sub cmdAdd_Click()
        '<EhHeader>
        On Error GoTo cmdAdd_Click_Err
        '</EhHeader>

        Dim RS As New ADODB.Recordset
        Dim sGUID As String
        Dim sTimeStamp As String
        Dim RSUpdater As ADODB.Recordset
        Dim lID As Long
            
100     DebugPrint ""
102     sGUID = GetGuid
104     sTimeStamp = Now()

106     If ComCategory.ListCount = 0 Or txtName = "" Or txtDesxription = "" Then
108         MsgBox "Please enter in all detail!"
        Else
    
            'mConn.Execute "INSERT INTO GeoBookMarks (Name, BmkrID, Description, X, Y, Z, MapName, [sGUID], [dTimeStamp]) VALUES ('" & txtName.Text & "'," & ComCategory.ItemData(ComCategory.ListIndex) & ",'" & IIf(Len(txtDesxription.Text) < 1, "N/A", txtDesxription.Text) & "'," & Replace(m_x, ",", ".") & "," & Replace(m_y, ",", ".") & "," & Replace(M_z, ",", ".") & ",'" & m_sMapName & "','" & sGUID & "',#" & sTimeStamp & "#)"
        
110         Set RSUpdater = New ADODB.Recordset

112         With RSUpdater

114             .Open "SELECT MAX(ID) FROM GeoBookMarks", mConn, adOpenDynamic, adLockBatchOptimistic

116             If Not .EOF And Not .Bof Then
118                 lID = CInt(IIf(IsNull(.Fields(0).Value), 0, .Fields(0).Value)) + 1
                Else
120                 lID = 1
                End If
                
122             .Close
124             .Open "SELECT * FROM GeoBookMarks", mConn, adOpenDynamic, adLockBatchOptimistic
                
126             .AddNew

                On Error Resume Next 'to deal with autonum
128              .Fields("ID").Value = lID
                On Error GoTo cmdAdd_Click_Err
130             .Fields("Name").Value = txtName.Text
132             .Fields("BmkrID").Value = ComCategory.ItemData(ComCategory.ListIndex)
134             .Fields("Description").Value = IIf(Len(txtDesxription.Text) < 1, "N/A", txtDesxription.Text)
136             .Fields("X").Value = Replace(m_x, ",", ".")
138             .Fields("Y").Value = Replace(m_y, ",", ".")
140             .Fields("Z").Value = Replace(M_z, ",", ".")
142             .Fields("MapName").Value = m_sMapName
144             .Fields("GUID1").Value = sGUID
146             .Fields("dTimeStamp").Value = sTimeStamp
148             .UpdateBatch adAffectCurrent
150             .Close

            End With

152         Set RSUpdater = Nothing

154         frmMain.SetNewSynchDBElement GetGuid, sGUID, "OASIS GeoMarks", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, "GeoBookMarks", False
        
        End If

156     Unload Me
        '<EhFooter>
        Exit Sub

cmdAdd_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.FrmAddBookMark.cmdAdd_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        Dim RS As New ADODB.Recordset
    
    If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
    
    
100     If mConn Is Nothing Then Exit Sub
    
102     RS.Open "SELECT Name, ID FROM GeoBookMarksCategories ORDER  BY Name", mConn
    
104     ComCategory.Clear
106     txtDesxription.Text = ""
108     txtName.Text = ""
    
110     If Not RS.Bof Then SafeMoveFirst RS
    
112     Do While Not RS.EOF
114         ComCategory.AddItem RS.Fields.Item("Name").Value
116         ComCategory.ItemData(ComCategory.ListCount - 1) = RS.Fields.Item("ID").Value
118         RS.MoveNext
        Loop
    
120     If ComCategory.ListCount > 0 Then ComCategory.ListIndex = 0
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.FrmAddBookMark.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
