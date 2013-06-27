VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim CliID As Long
Dim DF As CForm
Dim SWnd As Long
Public Event OnDataArrival(bData As String)
Attribute OnDataArrival.VB_Description = "Event fires when data is transmitted by the server process"
Public Event OnServerConnect(ByVal ChannelID As Long)
Attribute OnServerConnect.VB_Description = "Event fires when the server process connects to the active data channel"
Public Event OnServerDisconnect(ByVal ChannelID As Long)
Attribute OnServerDisconnect.VB_Description = "Event fires when the server process disconnects to the active data channel"
Public Event OnChannelRegistrationSuccess(ByVal ChannelID As Long)
Public Event OnChannelRegistrationFailure(ByVal ChannelID As Long, Reason As String)
Public Event OnVBInternalError(ByVal ErrCode As Long, ByVal ErrDesc As String)
Attribute OnVBInternalError.VB_Description = "Event fires when an unexpected internal error occurs"

Private Function doLog(sMess As String)


    SetProp DF.hwnd, "ClientLog", WindowHandle

     If gFrmComLog.Check1.Value = vbChecked Then
        If Len(gFrmComLog.txtComms.Text) > 1000 Then gFrmComLog.txtComms.Text = ""
        gFrmComLog.txtComms.Text = sMess & vbCrLf & gFrmComLog.txtComms.Text
     Else
        gFrmComLog.txtComms.Text = sMess
     End If
End Function

Public Sub RegisterDataChannel(ChannelID As Long, _
                        WindowHandle As Long)
Attribute RegisterDataChannel.VB_Description = "Registers a specified data channel and prepares it for data input"
        Dim Wnd As Long

        If ChannelID < 0 Then
            Err.Raise vbObjectError + 530, , "Cannot register Client with given ID"
            Exit Sub
        End If
        
        'doLog "OASIS Inter Comms Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID)
        
        Wnd = FindWindow(vbNullString, "OASIS Inter Comms Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID))
        
        If Wnd <> 0 Then
            If (GetProp(Wnd, "ClientWindow") <> 0) And (IsWindow(GetProp(Wnd, "ClientWindow")) = 0) Then
               'doLog Now() & " A previous client process had already registered [ID#]" & ChannelID & "]. OASIS Inter Comms will now attempt to redirect the channel to the current client"
               ' MsgBox "A previous client process had already registered [ID#]" & ChannelID & "]. OASIS Inter Comms will now attempt to redirect the channel to the current client", vbInformation Or vbOKOnly, "OASIS Inter Comms Critical Message"
                SendMessage Wnd, WM_CLOSE, 0, 0
            Else
                'doLog Now() & " Cannot Register Cliebnt With Given ID"
                RaiseEvent OnChannelRegistrationFailure(ChannelID, "Cannot register Client with given ID")
                Exit Sub
            End If
        End If

        If IsWindow(WindowHandle) = 0 Then
            'doLog Now() & " Invalid Window Handle Specified"
            RaiseEvent OnChannelRegistrationFailure(ChannelID, "Invalid window handle specified")
            Exit Sub
        End If

        On Local Error GoTo 500

        Set DF = New CForm
        Load DF
        DF.Visible = False
        DF.SetClient Me
        'doLog "Loaded Client Comms Win"
        DF.Caption = "OASIS Inter Comms Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID)
        CliID = ChannelID

        SetProp DF.hwnd, "ClientWindow", WindowHandle
        SetProp DF.hwnd, "ClientClass", ObjectPtr(Me)

        'Check Comms Server
        'doLog "Checking Server"
        SWnd = FindWindow(vbNullString, "OASIS Inter Comms Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID))

        If SWnd <> 0 Then
            SetProp SWnd, "Action", 2
            PostMessage SWnd, WM_SIZE, 0, 0
            DF.SetServerFormHandle SWnd
        End If

        RaiseEvent OnChannelRegistrationSuccess(ChannelID)
        'doLog Now() & " ChannelRegistration Success ID:" & ChannelID
        DF.SetID ChannelID
        Set GWin = GetGlobalWindow()
        GoTo 550
500     RaiseEvent OnVBInternalError(Err.Number, Err.Description)
        'doLog Now() & " Internal Error: " & Err.Description
        Err.Clear
550 End Sub

Public Function GetNextValidID() As Long
Attribute GetNextValidID.VB_Description = "Gets the next available valid Data Channel ID"
    Dim Wnd As Long, Iter As Long
    Iter = -1

    Do
        Iter = Iter + 1
        Wnd = FindWindow(vbNullString, "OASIS Inter Comms Data Channel ID[Client]:[Hidden Window]" & CStr(Iter))
    Loop Until Wnd = 0
    'doLog Now() & " GetNextValidID"
    GetNextValidID = Iter
End Function

Public Function IsChannelRegistered(ChannelID As Long) As Boolean
Attribute IsChannelRegistered.VB_Description = "Returns True if the specified channel is registered"
    Dim Wnd As Long
    Wnd = FindWindow(vbNullString, "OASIS Inter Comms Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID))

    If Wnd = 0 Then
        IsChannelRegistered = False
        'doLog Now() & "IsChannelRegistred = False"
    Else
        IsChannelRegistered = True
        'doLog Now() & "ISChannelRegistered = True"
    End If

End Function

Public Sub UnregisterChannel()
Attribute UnregisterChannel.VB_Description = "Unregisters the currently registered data channel"
    'doLog Now() & " Unreg Channel"
    On Error Resume Next
    Unload DF
    CliID = 0
    SetProp SWnd, "Action", 1
    PostMessage SWnd, WM_SIZE, 0, 0
    SWnd = 0
    Set DF = Nothing
70 End Sub

Friend Sub SetServerWnd(SHandle As Long)
    SWnd = SHandle
    DF.SetServerFormHandle SWnd
End Sub

Friend Sub RaiseDataEvent(bData As String)
        'doLog Now() & " Data Transfer: " & bData
        RaiseEvent OnDataArrival(bData)
        
        If IsWindow(SWnd) <> 0 Then
            PostMessage SWnd, WM_SIZE, 0, 0
        Else
            Dim GWin As GlobalWin
            Set GWin = GetGlobalWindow()
            GWin.NormaliseIcon
            'doLog Now() & " ON Server Disconnect"
            RaiseEvent OnServerDisconnect(CliID)
        End If

110 End Sub

Friend Sub RaiseConnectEvent()
    'doLog Now() & " OnServerConnect"
    RaiseEvent OnServerConnect(CliID)
End Sub
Friend Sub RaiseDisconnectEvent()
    SWnd = 0
    'doLog Now() & " OnServerDisConnect"
    RaiseEvent OnServerDisconnect(CliID)
End Sub

Friend Sub RaiseIntError(ErrNo As Long, _
                         ErrDesc As String)
    'doLog Now() & " Internal Error:" & ErrDesc & ErrNo
    RaiseEvent OnVBInternalError(ErrNo, ErrDesc)
End Sub

Public Sub SetCrashDetectInterval(Interval As Long)
On Error Resume Next

    If Interval < 0 Then Exit Sub
    If CFrm Is Nothing Then
    Else
        Dim SFrm As SForm
        Set SFrm = GetObject(GetProp("SForm", SWnd))

        If SFrm Is Nothing Then
        Else
            SFrm.SetInterval Interval
        End If
        
        DF.SetInterval Interval
    End If

End Sub
Private Sub Class_Initialize()
        On Error Resume Next
        Set gFrmComLog = New frmComLog
        Load gFrmComLog
        gFrmComLog.Visible = False
End Sub

Private Sub Class_Terminate()
        On Error Resume Next
        Unload gFrmComLog
        Set gFrmComLog = Nothing
End Sub
