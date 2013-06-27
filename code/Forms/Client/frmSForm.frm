VERSION 5.00
Begin VB.Form SForm 
   Caption         =   "Server Side Hidden Window"
   ClientHeight    =   870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   Icon            =   "frmSForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   870
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer CrashTimer 
      Interval        =   15000
      Left            =   30
      Top             =   3390
   End
End
Attribute VB_Name = "SForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ServerClass As IServer
Dim ID As Long

Private Sub CrashTimer_Timer()
On Error Resume Next
    If IsWindow(GetProp(Me.hwnd, "ServerWindow")) = 0 Then
        Dim CFormHandle As Long
        CFormHandle = FindWindow(vbNullString, "OASIS Inter Comms Data Channel ID[Client]:[Hidden Window]" & CStr(ID))
    
        SetProp Me.hwnd, "Busy", 1
        gFrmComLog.txtComms.Text = Now() & " A previous server process had connected to the [ID#" & ID & "]. OASIS Inter Comms will now attempt to terminate the connection"
        'MsgBox "A previous server process had connected to the [ID#" & ID & "]. OASIS Inter Comms will now attempt to terminate the connection", vbInformation Or vbOKOnly, "OASIS Inter Comms Critical Message"
        SetProp CFormHandle, "Action", 2
        PostMessage CFormHandle, WM_SIZE, 0, 0
        Set ServerClass = Nothing
        SendMessage Me.hwnd, WM_CLOSE, 0, 0
    End If
        
End Sub

Private Sub Form_Load()
On Error Resume Next
    IncrementServerCount
    SetProp Me.hwnd, "SForm", ObjectPtr(Me)
End Sub

Sub SetServerClass(SObj As IServer)
    Set ServerClass = SObj
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    If UnloadMode = vbAppWindows Then
        If GetProp(FindWindow(vbNullString, App.Title), "Confirm") = 0 Then
            Dim Choice As VbMsgBoxResult
            Dim Wnd As Long
            
'            Choice = MsgBox("Windows is attempting to shut down OASIS Inter Comms. Do you want to allow the shut down ?", vbCritical Or vbYesNo, "OASIS Inter Comms Critical Message")
'
'            If Choice = vbNo Then Cancel = True
'            If Choice = vbYes Then
                SetProp FindWindow(vbNullString, App.Title), "Confirm", 1
'            End If
        
        End If
    End If

End Sub

Private Sub Form_Resize()
        On Local Error GoTo 666

        Dim SAction As SRAction
        SAction = GetProp(Me.hwnd, "Action")

        If SAction = sacDefault Then
            If ServerClass Is Nothing Then
            Else
                ServerClass.RaiseSuccess
            End If
        End If

        If SAction = sacChannelClose Then
            ServerClass.RaiseChannelClose
        End If

        If SAction = sacChannelOpen Then
            ServerClass.RaiseChannelOpen
        End If

        SetProp Me.hwnd, "Action", 0
        GoTo 669
666     ServerClass.RaiseIntError Err.Number, Err.Description
        Err.Clear
669 End Sub

Private Sub Form_Unload(Cancel As Integer)
    DecrementServerCount
    Set ServerClass = Nothing
End Sub

Sub SetID(iID As Long)
    ID = iID
End Sub

Sub SetInterval(Interval As Long)
    CrashTimer.Interval = Interval
End Sub

