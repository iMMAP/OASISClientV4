VERSION 5.00
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Begin VB.Form frmChangePassword 
   BackColor       =   &H0050C0A4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User password change dialog"
   ClientHeight    =   2340
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5610
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton OKButton 
         Caption         =   "OK"
         Height          =   375
         Left            =   4290
         TabIndex        =   5
         Top             =   1410
         Width           =   1215
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4290
         TabIndex        =   6
         Top             =   1830
         Width           =   1215
      End
      Begin XpressEditorsLibCtl.dxMaskEdit txtOld 
         Height          =   315
         Left            =   90
         OleObjectBlob   =   "frmChangePassword.frx":6852
         TabIndex        =   2
         Top             =   480
         Width           =   4065
      End
      Begin XpressEditorsLibCtl.dxMaskEdit txtNew1 
         Height          =   315
         Left            =   90
         OleObjectBlob   =   "frmChangePassword.frx":68CD
         TabIndex        =   3
         Top             =   1170
         Width           =   4065
      End
      Begin XpressEditorsLibCtl.dxMaskEdit txtNew2 
         Height          =   315
         Left            =   90
         OleObjectBlob   =   "frmChangePassword.frx":6948
         TabIndex        =   4
         Top             =   1860
         Width           =   4065
      End
      Begin VB.Label Label2 
         BackColor       =   &H0050C0A4&
         BackStyle       =   0  'Transparent
         Caption         =   "Please confirm your NEW password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   1560
         Width           =   4125
      End
      Begin VB.Label Label1 
         BackColor       =   &H0050C0A4&
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter in your NEW password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Top             =   870
         Width           =   4125
      End
      Begin VB.Label lbl1 
         BackColor       =   &H0050C0A4&
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter in your OLD password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
        '<EhHeader>
        On Error GoTo CancelButton_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

CancelButton_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangePassword.CancelButton_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     txtNew1.EditStyle.BorderStyle = mbsSingle
102     txtNew2.EditStyle.BorderStyle = mbsSingle
104     txtOld.EditStyle.BorderStyle = mbsSingle
106     Call CheckFields
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangePassword.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub OKButton_Click()
        '<EhHeader>
        On Error GoTo OKButton_Click_Err
        '</EhHeader>

        Dim oRS As adodb.Recordset
        Dim sString As String
        Dim bRetVal As Boolean
    
100     If MsgBox("Are you sure you want to change your password?", vbYesNo, "Password change") = vbYes Then

102         sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT pwd FROM Users WHERE user = '" & g_sUserName & "'")
104         Set oRS = modServerComms.OpenSilentHttpCommsRS(sString, True)
106         oRS.Fields(0).Value = txtNew1
108         sString = g_sAppServerPath & "/oasis4.asp"
110         bRetVal = modServerComms.SaveSilentHttpCommsRS(oRS, sString, True)
        
112         If bRetVal Then
        
114             g_sUserPass = txtNew1
116             Set oRS = New adodb.Recordset
118             oRS.Open "SELECT pwd FROM Personnell WHERE Personnell_ID = 2", m_Cnn, adOpenDynamic, adLockBatchOptimistic
120             oRS.Fields(0).Value = g_sUserPass
122             oRS.UpdateBatch adAffectCurrent
124             Set oRS = Nothing
126             MsgBox "Password changed"
128             Unload Me
            Else
130             Set oRS = Nothing
132             MsgBox "Password change failed!"
            End If

        End If

        '<EhFooter>
        Exit Sub

OKButton_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangePassword.OKButton_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CheckFields()
        '<EhHeader>
        On Error GoTo CheckFields_Err
        '</EhHeader>

        Dim bGood As Boolean
    
100     If Not (txtNew1 = "") And (txtNew1 = txtNew2) Then
 
102         txtNew1.EditStyle.BorderColor = vbGreen
104         txtNew2.EditStyle.BorderColor = vbGreen
106         bGood = True

        Else
 
108         txtNew1.EditStyle.BorderColor = vbRed
110         txtNew2.EditStyle.BorderColor = vbRed

        End If
    
112     If txtOld = g_sUserPass Then
    
114         txtOld.EditStyle.BorderColor = vbGreen
116         txtOld.EditStyle.BorderColor = vbGreen

        Else
 
118         txtOld.EditStyle.BorderColor = vbRed
120         txtOld.EditStyle.BorderColor = vbRed

122         If bGood Then bGood = False

        End If
    
124     If bGood Then
126         OKButton.Visible = True
        Else
128         OKButton.Visible = False
        End If

        '<EhFooter>
        Exit Sub

CheckFields_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangePassword.CheckFields " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtNew1_Change()
        '<EhHeader>
        On Error GoTo txtNew1_Change_Err
        '</EhHeader>

100     Call CheckFields

        '<EhFooter>
        Exit Sub

txtNew1_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangePassword.txtNew1_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtNew2_Change()
        '<EhHeader>
        On Error GoTo txtNew2_Change_Err
        '</EhHeader>
100     Call CheckFields
        '<EhFooter>
        Exit Sub

txtNew2_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangePassword.txtNew2_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtNew2_KeyDown(KeyCode As Integer, _
                            Shift As Integer)
        '<EhHeader>
        On Error GoTo txtNew2_KeyDown_Err
        '</EhHeader>

100     If KeyCode = 13 Then
102         Call OKButton_Click
        End If

        '<EhFooter>
        Exit Sub

txtNew2_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangePassword.txtNew2_KeyDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtOld_Change()
        '<EhHeader>
        On Error GoTo txtOld_Change_Err
        '</EhHeader>

100     Call CheckFields

        '<EhFooter>
        Exit Sub

txtOld_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangePassword.txtOld_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtOld_KeyDown(KeyCode As Integer, _
                           Shift As Integer)
        '<EhHeader>
        On Error GoTo txtOld_KeyDown_Err
        '</EhHeader>

100     If KeyCode = 13 Then
102         Call OKButton_Click
        End If

        '<EhFooter>
        Exit Sub

txtOld_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangePassword.txtOld_KeyDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtNew1_KeyDown(KeyCode As Integer, _
                            Shift As Integer)
        '<EhHeader>
        On Error GoTo txtNew1_KeyDown_Err
        '</EhHeader>

100     If KeyCode = 13 Then
102         Call OKButton_Click
        End If

        '<EhFooter>
        Exit Sub

txtNew1_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangePassword.txtNew1_KeyDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
