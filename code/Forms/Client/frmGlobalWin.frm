VERSION 5.00
Begin VB.Form GlobalWin 
   AutoRedraw      =   -1  'True
   Caption         =   "OASIS Inter Comm Connection Manager"
   ClientHeight    =   1185
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   79
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Active 
      Height          =   480
      Left            =   1410
      Picture         =   "frmGlobalWin.frx":0000
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Inactive 
      Height          =   480
      Left            =   675
      Picture         =   "frmGlobalWin.frx":0442
      Top             =   2145
      Width           =   480
   End
   Begin VB.Menu TrayMenu 
      Caption         =   "&Tray Menu"
      Begin VB.Menu Terminate 
         Caption         =   "&Terminate OASIS Inter Comms...."
      End
      Begin VB.Menu SEP2 
         Caption         =   "-"
      End
      Begin VB.Menu CStat 
         Caption         =   "&Connection Status...."
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu CLog 
         Caption         =   "&Show Log...."
      End
   End
End
Attribute VB_Name = "GlobalWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tray As NOTIFYICONDATA
Dim Busy As Boolean

Private Sub CLog_Click()
    'If Not gFrmComLog Is Nothing Then gFrmComLog.Visible = True
    
    Dim GLog As frmComLog
    
    On Error Resume Next
    
    Set GLog = GetLogWindow()
    
    If Not GLog Is Nothing Then GLog.Visible = True
End Sub

Private Sub CStat_Click()
    MsgBox "Total Synch Processes Registered: " & GetProp(Me.hwnd, "ServerCount") & Chr$(13) & "Total Client Connections Registered: " & GetProp(Me.hwnd, "ClientCount") & Chr$(13), vbInformation, "OASIS Inter Comms Connection Status"
End Sub

Private Sub Form_Load()
On Error Resume Next
    SetProp Me.hwnd, "GlobalWin", ObjectPtr(Me)

    With Tray
        .cbSize = Len(Tray)
        .hwnd = Me.hwnd
        .hIcon = Inactive.Picture
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .szTip = "OASIS Synch Communicator" & vbNullChar
        .uCallBackMessage = WM_MOUSEMOVE
    End With

    Shell_NotifyIcon NIM_ADD, Tray
End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
    Dim GLog As frmComLog
    
    On Error Resume Next
    
    Select Case X

        Case WM_RBUTTONUP
            Set GLog = GetLogWindow()
    
            If GLog Is Nothing Then
                CLog.Visible = False
            Else
                CLog.Visible = True
            End If
            
            Me.PopupMenu Me.TrayMenu
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, Tray
End Sub

Public Sub NormaliseIcon()
    On Error Resume Next
    
    With Tray
        .cbSize = Len(Tray)
        .hwnd = Me.hwnd
        .hIcon = Inactive.Picture
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .szTip = "OASIS Synch Communicator" & vbNullChar
        .uCallBackMessage = WM_MOUSEMOVE
    End With

    Shell_NotifyIcon NIM_MODIFY, Tray
    Busy = False
End Sub

Public Sub HighlightIcon()

    With Tray
        .cbSize = Len(Tray)
        .hwnd = Me.hwnd
        .hIcon = Active.Picture
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .szTip = "OASIS Synch Communicator" & vbNullChar
        .uCallBackMessage = WM_MOUSEMOVE
    End With

    Shell_NotifyIcon NIM_MODIFY, Tray
    Busy = True
End Sub

Public Function GetBusyFlag() As Boolean
    GetBusyFlag = GWin
End Function

Private Sub Terminate_Click()
    Dim Choice As VbMsgBoxResult
    Choice = MsgBox("Terminating OASIS Synch Communicator. This will abort all Synchronisation Communication. Do you still want to terminate?", vbYesNo Or vbDefaultButton2 Or vbCritical, "OASIS Inter Comms Critical Message")

    If Choice = vbYes Then End
End Sub
