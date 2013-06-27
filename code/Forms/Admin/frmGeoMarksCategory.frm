VERSION 5.00
Begin VB.Form frmGeoMarksCategory 
   BackColor       =   &H0050C0A4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geo Bookmarks Category"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3360
   Icon            =   "frmGeoMarksCategory.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstGeoMarkGategory 
      Height          =   2595
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   2820
      Width           =   3315
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   285
      Left            =   2580
      TabIndex        =   3
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtDesc 
      Height          =   945
      Left            =   60
      TabIndex        =   0
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label lblName 
      BackColor       =   &H0050C0A4&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   2610
      Width           =   1215
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H0050C0A4&
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   3120
      Width           =   1605
   End
End
Attribute VB_Name = "frmGeoMarksCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RSGeoMarkCategories As ADODB.Recordset
Private RSLocalUserGroups As ADODB.Recordset
Public Event RefreshCategories()

Private Sub cmdExit_Click()
        '<EhHeader>
        On Error GoTo cmdExit_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

cmdExit_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksCategory.cmdExit_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CheckEdit() As Boolean
        '<EhHeader>
        On Error GoTo CheckEdit_Err
        '</EhHeader>
        Dim bReturnValue As Boolean
        Dim sString As String

100     With RSGeoMarkCategories
102         .Filter = adFilterPendingRecords
    
104         If Not .Bof And Not .EOF Then
106             If MsgBox("Do you wish to save your changes?", vbYesNo, "Confirm Save") = vbYes Then

108                 sString = WebSite & "Oasis.asp"
110                 bReturnValue = m_frmOASISProgress.SaveHttpCommsRS(RSGeoMarkCategories, sString, True)
                    
112                 If bReturnValue Then
114                     MsgBox "Data saved to server"
                    Else
116                     MsgBox "Saving to server failed!"
                    End If
                    
                End If

118             CheckEdit = True
            End If

        End With

        '<EhFooter>
        Exit Function

CheckEdit_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksCategory.CheckEdit " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdNew_Click()
        '<EhHeader>
        On Error GoTo cmdNew_Click_Err
        '</EhHeader>

100     If Not IsNull(Me.txtName) And Len(txtName) > 0 Then
102         RSGeoMarkCategories.AddNew
104         RSGeoMarkCategories.fields("Name").Value = Me.txtName
106         RSGeoMarkCategories.fields("Description").Value = Me.txtDesc
            RSGeoMarkCategories.fields("sGUID").Value = GUIDGen

108         LoadCategories
110         Me.lstGeoMarkGategory.Text = txtName
        End If

        '<EhFooter>
        Exit Sub

cmdNew_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksCategory.cmdNew_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        Dim sString As String

100     Set RSGeoMarkCategories = New ADODB.Recordset

102     sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & RSLocalUserGroups!Name & "GeoBookMarksCategories ORDER BY Name")
104     Set RSGeoMarkCategories = m_frmOASISProgress.OpenHttpCommsRS(sString, True)

106     If RSGeoMarkCategories.State = adStateClosed Then
108         MsgBox "Something Shifty went on @ the server....!"
            Exit Sub
        End If

110     Set txtDesc.DataSource = RSGeoMarkCategories
112     Set txtName.DataSource = RSGeoMarkCategories

114     LoadCategories

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksCategory.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadCategories()
        '<EhHeader>
        On Error GoTo LoadCategories_Err
        '</EhHeader>

100     With RSGeoMarkCategories
102         lstGeoMarkGategory.Clear

104         If Not .EOF Or Not .Bof Then
106             .MoveFirst
            
108             Do While Not .EOF
110                 lstGeoMarkGategory.AddItem .fields.Item("Name").Value
                    'lstGeoMarkGategory.ItemData(lstGeoMarkGategory.ListCount - 1) = CLng(.fields.Item("ID").Value)
112                 .MoveNext
                Loop
            
114             lstGeoMarkGategory.ListIndex = 0
            
            End If
    
        End With
    
        '<EhFooter>
        Exit Sub

LoadCategories_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksCategory.LoadCategories " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub setUserGroupsRS(ByVal PassedRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo setUserGroupsRS_Err
        '</EhHeader>
100     Set RSLocalUserGroups = PassedRS
    
        '<EhFooter>
        Exit Sub

setUserGroupsRS_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksCategory.setUserGroupsRS " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     CheckEdit
102     RaiseEvent RefreshCategories
104     Unload Me
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksCategory.Form_Unload " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

