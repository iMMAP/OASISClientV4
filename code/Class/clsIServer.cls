VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event OnTransmissionSuccess()
Attribute OnTransmissionSuccess.VB_Description = "Event fires when an attempt to transmit data to the client app succeeds"
Public Event OnTransmissionFailure(Reason As String)
Attribute OnTransmissionFailure.VB_Description = "Event fires when an attempt to transmit data to the client app fails"
Public Event OnConnectionWait()
Attribute OnConnectionWait.VB_Description = "Event fires when the server instantiates a data channel and is is beginning to wait for the client side interface to be instantiated by client process/thread"
Public Event OnConnectionSuccess()
Attribute OnConnectionSuccess.VB_Description = "Event fires when an attempt to connect to a data channel succeeds"
Public Event OnConnectionFailure(Reason As String)
Attribute OnConnectionFailure.VB_Description = "Event fires when an attempt to connect to a data channel fails"
Public Event OnChannelReOpen(ByVal ChannelID As Long)
Attribute OnChannelReOpen.VB_Description = "Event fires when the client reopens  the data channel associated with the given instance of InterCommVB's server side interface"
Public Event OnChannelOpen(ByVal ChannelID As Long)
Attribute OnChannelOpen.VB_Description = "Event fires when the client opens the associated data channel "
Public Event OnChannelClose(ByVal ChannelID As Long)
Attribute OnChannelClose.VB_Description = "Event fires when the client closes the data channel associated with the given instance of InterCommVB's server side interface"
Public Event OnVBInternalError(ByVal ErrCode As Long, ByVal ErrDesc As String)
Attribute OnVBInternalError.VB_Description = "Event fires when an unexpected internal error occurs"
Dim CWnd As Long, CFrm As CForm, CID As Long
Dim ServerInit As Boolean, GWin As GlobalWin, GLog As frmComLog
Dim SFrm As SForm, ClientClass As IClient

Private Function doLog(sMess As String)
    
    If GLog Is Nothing Then Set GLog = GetLogWindow()

    If Not GLog Is Nothing Then
        If GLog.Check1.Value = vbChecked Then
            If Len(GLog.txtComms.Text) > 7000 Then GLog.txtComms.Text = ""
            GLog.txtComms.Text = sMess & vbCrLf & GLog.txtComms.Text
        Else
            GLog.txtComms.Text = sMess
        End If
    End If
     
End Function

Public Sub ConnectToDataChannel(ChannelID As Long, _
                         WindowHandle As Long, _
                         Optional IgnoreMissingClient As Boolean = False)
Attribute ConnectToDataChannel.VB_Description = "Connects to the specified data channel"
    On Local Error GoTo 10

    If CWnd <> 0 Then
        Call DisconnectFromChannel
    End If

    If IsWindow(WindowHandle) = 0 Then
        doLog Now() & " Invalid Handle"
        RaiseEvent OnConnectionFailure("Invalid Window handle specified")
        GoTo 150
        Exit Sub
    End If

    Dim Wnd As Long
    Dim DFPtr As Long, DF As CForm, STwnd As Long
    
    STwnd = FindWindow(vbNullString, "OASIS Inter Comms Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID))

    If STwnd <> 0 Then
        If (GetProp(STwnd, "ServerWindow") <> 0) And (IsWindow(GetProp(STwnd, "ServerWindow")) = 0) Then
            doLog Now() & " A previous server process had connected to the requested data channel [ID#" & ChannelID & "]. OASIS Inter Comms will now attempt to redirect the connection to this server"
            'MsgBox "A previous server process had connected to the requested data channel [ID#" & ChannelID & "]. OASIS Inter Comms will now attempt to redirect the connection to this server", vbInformation Or vbOKOnly, "OASIS Inter Comms Critical Message"
            SendMessage STwnd, WM_CLOSE, 0, 0
        Else
            RaiseEvent OnConnectionFailure("Another process or thread has already connected to this channel")
            
            doLog Now() & " Another process or thread has already connected to this channel"
            
            GoTo 150
            Exit Sub
        End If
    End If
    
    doLog Now() & " Connect To Channel... "
    Wnd = FindWindow(vbNullString, "OASIS Inter Comms Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID))

    If (Wnd = 0) And (IgnoreMissingClient = True) Then
        ServerInit = True
        doLog Now() & " Ignore Client And Connect"
        Call IgnoreClientAndConnect(WindowHandle, ChannelID)
        GoTo 150
    End If

    ServerInit = False

    If Wnd = 0 Then
        RaiseEvent OnConnectionFailure("Unable to find data channel")
        doLog Now() & " Unable to find Data Channel"
        GoTo 150
        Exit Sub
    End If

    If GetProp(Wnd, "ClientClass") = 0 Then
        doLog Now() & " Unable to get pointer to Client COM interface"
        RaiseEvent OnConnectionFailure("Unable to get pointer to client side COM interface[IClient]")
        GoTo 150
        Exit Sub
    End If

    If GetProp(Wnd, "Busy") = 1 Then
        doLog Now() & " Comms Channel is Busy"
        RaiseEvent OnConnectionFailure("This communication channel is busy - A previous request is pending")
        GoTo 150
        Exit Sub
    End If

    CWnd = Wnd

    DFPtr = GetProp(Wnd, "CForm")
    Set CFrm = GetObj(DFPtr)

    doLog Now() & " LOADING SERVER FORM"
    Set SFrm = New SForm
    Load SFrm
    SFrm.Visible = False
    SetProp SFrm.hwnd, "ServerWindow", WindowHandle
    SFrm.Caption = "OASIS Inter Comms Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID)
    SFrm.SetServerClass Me
    
    doLog Now() & " Finding the Client"
    Set ClientClass = GetObj(GetProp(Wnd, "ClientClass"))
    ClientClass.SetServerWnd SFrm.hwnd
    
    doLog Now() & " Success Finding Client"
    RaiseEvent OnConnectionSuccess

    SetProp CWnd, "Action", 1
    PostMessage CWnd, WM_SIZE, 0, 0
    SFrm.SetID ChannelID
    doLog Now() & " Getting Global DDE"
    Set GWin = GetGlobalWindow()
    
    GoTo 20
10  RaiseEvent OnVBInternalError(Err.Number, Err.Description)
    Err.Clear

20  CID = ChannelID
150 End Sub

Public Sub TransmitToClient(bData As String)
Attribute TransmitToClient.VB_Description = "Transmits the specified data to the client process"

        On Local Error GoTo 30

        If (CWnd = 0) And (CID <> 0) Then
            AttemptReConnect CID
            
            doLog Now() & " Trying To Connect for data..."
            
            If CWnd <> 0 Then
                
                doLog Now() & " Data to transmit:" & bData
                
                CFrm.SetData bData
                SetProp CWnd, "Action", 0
                PostMessage CWnd, WM_SIZE, 0, 0
            Else
                GoTo 250
            End If
        End If

        If GetProp(CWnd, "Busy") = 1 Then
            doLog Now() & " Busy... Previous Mess Pending..."
            RaiseEvent OnTransmissionFailure("This communication channel is busy - A previous request is pending")
            GoTo 50
            Exit Sub
        End If

250     If (IsWindow(CWnd) = 0) Or (IsWindow(GetProp(CWnd, "ClientWindow")) = 0) Then
            Set CFrm = Nothing
            
            doLog Now() & " Data Channel Is not Open"
            RaiseEvent OnTransmissionFailure("The Data channel is not open")
    
            If (IsWindow(GetProp(CWnd, "ClientWindow")) = 0) Then SendMessage CWnd, WM_CLOSE, 0, 0
            CWnd = 0
            Set ClientClass = Nothing
            GoTo 50
            Exit Sub
        End If
    
    If GWin.GetBusyFlag = False Then GWin.HighlightIcon
    
    doLog Now() & " DoSetData: " & bData
    CFrm.SetData bData
    SetProp CWnd, "Action", 0
    PostMessage CWnd, WM_SIZE, 0, 0
    GoTo 50
30  RaiseEvent OnVBInternalError(Err.Number, Err.Description)
    Err.Clear
50 End Sub

Public Function IsChannelRegistered(ChannelID As Long) As Boolean
Attribute IsChannelRegistered.VB_Description = "Returns true if the specified data channel is registered"
    Dim Wnd As Long
    Wnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID))

    If Wnd = 0 Then
        IsChannelRegistered = False
    Else
        IsChannelRegistered = True
    End If

End Function

Public Sub DisconnectFromChannel()
Attribute DisconnectFromChannel.VB_Description = "DisConnects from  the specified data channel"
    Set CFrm = Nothing
    Set GWin = Nothing
    Set GLog = Nothing
    doLog Now() & " Disconnecting..."

    If SFrm Is Nothing Then
    Else
        Unload SFrm
        Set SFrm = Nothing
    End If

    Set ClientClass = Nothing

    If GetProp(CWnd, "Busy") = 0 Then
        SetProp CWnd, "Action", 2
        PostMessage CWnd, WM_SIZE, 0, 0
    End If

    CID = 0
    CWnd = 0
    
End Sub

Public Function IsConnected(ChannelID As Long) As Boolean
Attribute IsConnected.VB_Description = "Returns true if a server app has connected to a given data channel"
    Dim STwnd As Long
    STwnd = FindWindow(vbNullString, "OASIS Inter Comms Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID))

    If STwnd = 0 Then
        IsConnected = False
        doLog Now() & " IsConnected=False"
    Else
        doLog Now() & " IsConnected=True"
        IsConnected = True
        
    End If

End Function

Friend Sub RaiseSuccess()
    doLog Now() & " Transmission Success"
    RaiseEvent OnTransmissionSuccess

    If GWin.GetBusyFlag = False Then GWin.NormaliseIcon
End Sub

Friend Sub RaiseChannelClose()

    Set CFrm = Nothing
    CWnd = 0
    
    doLog Now() & " Channel Close"
    
    Set ClientClass = Nothing
    RaiseEvent OnChannelClose(CID)
End Sub

Friend Sub RaiseChannelOpen()
    AttemptReConnect CID

    doLog Now() & " On Channel Open"

    If ServerInit = False Then
        RaiseEvent OnChannelReOpen(CID)
    Else
        RaiseEvent OnChannelOpen(CID)
        ServerInit = False
    End If

End Sub

Friend Sub RaiseIntError(ErrNo As Long, _
                         ErrDesc As String)
                         doLog Now() & " ERROR:" & ErrDesc & ErrNo
    RaiseEvent OnVBInternalError(ErrNo, ErrDesc)
End Sub

Friend Sub AttemptReConnect(ChannelID As Long)
        On Local Error GoTo 400

        Dim Wnd As Long
        Dim DFPtr As Long, DF As CForm
        
        doLog Now() & " Attempt Reconnect"
        Wnd = FindWindow(vbNullString, "OASIS Inter Comms Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID))
        SWnd = FindWindow(vbNullString, "OASIS Inter Comms Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID))

        If Wnd = 0 Then
            GoTo 350
            Exit Sub
        End If

        If GetProp(Wnd, "ClientClass") = 0 Then
            GoTo 350
            Exit Sub
        End If

        If GetProp(Wnd, "Busy") = 1 Then
            GoTo 350
            Exit Sub
        End If

        CWnd = Wnd
        DFPtr = GetProp(Wnd, "CForm")

        Set CFrm = GetObj(DFPtr)
        Set ClientClass = GetObj(GetProp(Wnd, "ClientClass"))

        ClientClass.SetServerWnd SFrm.hwnd
        CID = ChannelID

        GoTo 350
400     RaiseEvent OnVBInternalError(Err.Number, Err.Description)
        Err.Clear
350 End Sub

Friend Sub IgnoreClientAndConnect(WindowHandle As Long, _
                                  ChannelID As Long)
        On Local Error GoTo 1000

        If CWnd <> 0 Then
            Call DisconnectFromChannel
        End If

        If IsWindow(WindowHandle) = 0 Then
            RaiseEvent OnConnectionFailure("Invalid Window handle specified")
            GoTo 15000
            Exit Sub
        End If

        Dim Wnd As Long
        Dim DFPtr As Long, DF As CForm, STwnd As Long
        
        STwnd = FindWindow(vbNullString, "OASIS Inter Comms Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID))

        If STwnd <> 0 Then
            If (GetProp(STwnd, "ServerWindow") <> 0) And (IsWindow(GetProp(STwnd, "ServerWindow")) = 0) Then
                doLog Now() & " A previous server process had connected [ID#" & ChannelID & "]. OASIS Inter Comms will now attempt to redirect the connection to this server"
               ' MsgBox "A previous server process had connected [ID#" & ChannelID & "]. OASIS Inter Comms will now attempt to redirect the connection to this server", vbInformation Or vbOKOnly, "OASIS Inter Comms Critical Message"
                SendMessage STwnd, WM_CLOSE, 0, 0
            Else
                doLog Now() & " Another Process is already Connected to This Channel"
                RaiseEvent OnConnectionFailure("Another process or thread has already connected to this channel")
                GoTo 15000
                Exit Sub
            End If
        End If
    
        CWnd = 0

        Set SFrm = New SForm
        Load SFrm
        SFrm.Visible = False
        SetProp SFrm.hwnd, "ServerWindow", WindowHandle
        SFrm.Caption = "OASIS Inter Comms Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID)
        SFrm.SetServerClass Me
        SFrm.SetID ChannelID
        
        doLog Now() & " On Connect Wait"
        
        RaiseEvent OnConnectionWait
        Set GWin = GetGlobalWindow()
        
        GoTo 2000
1000    RaiseEvent OnVBInternalError(Err.Number, Err.Description)
        Err.Clear

2000    CID = ChannelID
15000 End Sub

Public Sub SetCrashDetectInterval(Interval As Long)

    If Interval < 0 Then Exit Sub
    If SFrm Is Nothing Then
    Else
        SFrm.SetInterval Interval
        CFrm.SetInterval Interval
    End If

End Sub

'Private Sub Class_Initialize()
'        Set gFrmComLog = New frmComLog
'        Load gFrmComLog
'        gFrmComLog.Visible = False
'End Sub
'
'Private Sub Class_Terminate()
'        On Error Resume Next
'        Unload gFrmComLog
'        Set gFrmComLog = Nothing
'End Sub
