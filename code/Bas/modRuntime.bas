Attribute VB_Name = "modRuntime"

Public Declare Function PostMessage _
               Lib "user32" _
               Alias "PostMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     ByVal lParam As Long) As Long
Public Declare Function IsWindow _
               Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetProp _
               Lib "user32" _
               Alias "SetPropA" (ByVal hwnd As Long, _
                                 ByVal lpString As String, _
                                 ByVal hData As Long) As Long
Public Declare Function RemoveProp _
               Lib "user32" _
               Alias "RemovePropA" (ByVal hwnd As Long, _
                                    ByVal lpString As String) As Long
Public Declare Function GetProp _
               Lib "user32" _
               Alias "GetPropA" (ByVal hwnd As Long, _
                                 ByVal lpString As String) As Long
Public Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (Destination As Any, _
                                      Source As Any, _
                                      ByVal Length As Long)
Public Declare Function FindWindow _
               Lib "user32" _
               Alias "FindWindowA" (ByVal lpClassName As String, _
                                    ByVal lpWindowName As String) As Long
Public Declare Function TerminateThread _
               Lib "kernel32" (ByVal hThread As Long, _
                               ByVal dwExitCode As Long) As Long
Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long
Public Declare Function Shell_NotifyIcon _
               Lib "shell32" _
               Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                          pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow _
               Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCurrentProcess _
               Lib "kernel32" () As Long
Public Declare Function GetCurrentThread _
               Lib "kernel32" () As Long
Public Declare Function DuplicateHandle _
               Lib "kernel32" (ByVal hSourceProcessHandle As Long, _
                               ByVal hSourceHandle As Long, _
                               ByVal hTargetProcessHandle As Long, _
                               lpTargetHandle As Long, _
                               ByVal dwDesiredAccess As Long, _
                               ByVal bInheritHandle As Long, _
                               ByVal dwOptions As Long) As Long
Public Declare Function SetThreadPriority _
               Lib "kernel32" (ByVal hThread As Long, _
                               ByVal nPriority As Long) As Long

Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const THREAD_BASE_PRIORITY_MIN = -2

Public Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Public Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN

Public Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Public Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)

'//UDT required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
    cbSize As Long             '//size of this UDT
    hwnd As Long               '//handle of the app
    uId As Long                '//unused (set to vbNull)
    uFlags As Long             '//Flags needed for actions
    uCallBackMessage As Long   '//WM we are going to subclass
    hIcon As Long              '//Icon we're going to use for the systray
    szTip As String * 64       '//ToolTip for the mouse_over of the icon.
End Type

'//Constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0             '//Flag : "ALL NEW nid"
Public Const NIM_MODIFY = &H1          '//Flag : "ONLY MODIFYING nid"
Public Const NIM_DELETE = &H2          '//Flag : "DELETE THE CURRENT nid"
Public Const NIF_MESSAGE = &H1         '//Flag : "Message in nid is valid"
Public Const NIF_ICON = &H2            '//Flag : "Icon in nid is valid"
Public Const NIF_TIP = &H4             '//Flag : "Tip in nid is valid"
Public Const WM_MOUSEMOVE = &H200      '//This is our CallBack Message
Public Const WM_LBUTTONDOWN = &H201    '//LButton down
Public Const WM_LBUTTONUP = &H202      '//LButton up
Public Const WM_LBUTTONDBLCLK = &H203  '//LDouble-click
Public Const WM_RBUTTONDOWN = &H204    '//RButton down
Public Const WM_RBUTTONUP = &H205      '//RButton up
Public Const WM_RBUTTONDBLCLK = &H206  '//RDouble-click

Public nid As NOTIFYICONDATA       '//global UDT for the systray function

Public Const WM_CLOSE = &H10

Public Const WM_SIZE = &H5
Public Enum CRAction
    acDefault = 0
    acConnectEvent = 1
    acDisConnectEvent = 2
End Enum

Public Enum SRAction
    sacDefault = 0
    sacChannelClose = 1
    sacChannelOpen = 2
End Enum

Public gFrmComLog As frmComLog

Public Function GetObj(Ptr As Long)
    'Retrieves an Object from its pointer
    Dim TObj As Object
    CopyMemory TObj, Ptr, 4
    Set GetObj = TObj
    CopyMemory TObj, 0&, 4
End Function

Public Function ObjectPtr(Obj As Object)
    'Returns a pointer to an object
    Dim lpObj As Long
    CopyMemory lpObj, Obj, 4
    ObjectPtr = lpObj
End Function

Public Sub cLoadGlobalWin()
    Dim Wnd As Long
    
    On Error Resume Next
    
    Wnd = FindWindow(vbNullString, "OASIS Inter Comm Connection Manager")

    If Wnd = 0 Then
        Dim GFrm As New GlobalWin
        Load GFrm
        GFrm.Visible = False
        Set GFrm = Nothing
    End If


    Wnd = FindWindow(vbNullString, "OASIS Comms Log")

    If Wnd = 0 Then
        Dim GFlOG As New frmComLog
        Load GFlOG
        GFlOG.Visible = False
        Set GFlOG = Nothing
    End If

End Sub

Public Sub SetGlobalProp(PropName As String, _
                  PVal As Long)
    Dim Wnd As Long
    On Error Resume Next
    Wnd = FindWindow(vbNullString, "OASIS Inter Comm Connection Manager")
    SetProp Wnd, PropName, PVal
End Sub

Public Function GetGlobalProp(PropName As String) As Long
    Dim Wnd As Long
    On Error Resume Next
    Wnd = FindWindow(vbNullString, "OASIS Inter Comm Connection Manager")
    GetGlobalProp = GetProp(Wnd, PropName)
End Function

Public Sub RemoveGlobalProp(PropName As String)
    Dim Wnd As Long
    On Error Resume Next
    Wnd = FindWindow(vbNullString, "OASIS Inter Comm Connection Manager")
    RemoveProp Wnd, PropName
End Sub

Public Sub IncrementClientCount()
    Dim Wnd As Long
    On Error Resume Next
    Wnd = FindWindow(vbNullString, "OASIS Inter Comm Connection Manager")
    SetProp Wnd, "ClientCount", GetGlobalProp("ClientCount") + 1
End Sub

Public Sub DecrementClientCount()
    Dim Wnd As Long
    On Error Resume Next
    Wnd = FindWindow(vbNullString, "OASIS Inter Comm Connection Manager")
    SetProp Wnd, "ClientCount", GetGlobalProp("ClientCount") - 1
    Call ChkTerminate
End Sub

Public Sub IncrementServerCount()
    Dim Wnd As Long
    On Error Resume Next
    Wnd = FindWindow(vbNullString, "OASIS Inter Comm Connection Manager")
    SetProp Wnd, "ServerCount", GetGlobalProp("ServerCount") + 1
End Sub

Public Sub DecrementServerCount()
    Dim Wnd As Long
    On Error Resume Next
    Wnd = FindWindow(vbNullString, "OASIS Inter Comm Connection Manager")
    SetProp Wnd, "ServerCount", GetGlobalProp("ServerCount") - 1
    Call ChkTerminate
End Sub

Public Sub ChkTerminate()

On Error Resume Next

    If GetGlobalProp("ServerCount") = 0 And GetGlobalProp("ClientCount") = 0 Then
        Dim Wnd As Long
        Wnd = FindWindow(vbNullString, "OASIS Inter Comm Connection Manager")
        SendMessage Wnd, WM_CLOSE, 0, 0
    End If

End Sub

Public Function GetGlobalWindow() As GlobalWin
    Dim Wnd As Long
    On Error Resume Next
    Wnd = FindWindow(vbNullString, "OASIS Inter Comm Connection Manager")
    Set GetGlobalWindow = GetObj(GetProp(Wnd, "GlobalWin"))
End Function

Public Function GetLogWindow() As frmComLog
    Dim Wnd As Long
    Wnd = FindWindow(vbNullString, "OASIS Comms Log")
    Set GetLogWindow = GetObj(GetProp(Wnd, "frmComLog"))
End Function

Public Sub Main()
On Error Resume Next
    cLoadGlobalWin
End Sub

