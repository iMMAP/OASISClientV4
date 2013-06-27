VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cClipboardViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Private Declare Function SetClipboardViewer Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function ChangeClipboardChain Lib "USER32" (ByVal hWnd As Long, ByVal hWndNext As Long) As Long
Private Declare Function SendMessageByString Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageByLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_CHANGECBCHAIN = &H30D
Private Const WM_DRAWCLIPBOARD = &H308

Private m_hWnd As Long
Private m_hWndNextViewer As Long
Private m_bInClipboardChangeNotification As Boolean
Private m_emr As EMsgResponse

Public Event ClipboardChanged()

Public Sub InitClipboardChangeNotification(hWndA As Long)
    StopClipboardChangeNotification
    m_hWnd = hWndA
    If (m_hWnd <> 0) Then
        ' Attach Clipboard 'viewer' notification messages
        AttachMessage Me, m_hWnd, WM_CHANGECBCHAIN
        AttachMessage Me, m_hWnd, WM_DRAWCLIPBOARD
        ' Place me in the clipboard viewer notification chain:
        m_hWndNextViewer = SetClipboardViewer(m_hWnd)
        m_bInClipboardChangeNotification = True
    End If
End Sub
Public Sub StopClipboardChangeNotification()
    If (m_bInClipboardChangeNotification) Then
        If (m_hWnd <> 0) Then
            ' Take myself out of the clipboard chain:
            ChangeClipboardChain m_hWnd, m_hWndNextViewer
            ' Stop subclassing for clipboard messages:
            DetachMessage Me, m_hWnd, WM_CHANGECBCHAIN
            DetachMessage Me, m_hWnd, WM_DRAWCLIPBOARD
        End If
    End If
    m_bInClipboardChangeNotification = False
End Sub
Private Sub Class_Terminate()
   StopClipboardChangeNotification
End Sub

Private Function ISubclass_WindowProc(ByVal hWnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
    Select Case iMsg
    Case WM_CHANGECBCHAIN
        If (wParam = m_hWndNextViewer) Then
            ' If the next viewer window is closing, repair the chain:
            m_hWndNextViewer = lParam
        ElseIf (m_hWndNextViewer <> 0) Then
            ' Otherwise if there is a next window, pass the message on:
            SendMessageByLong m_hWndNextViewer, iMsg, wParam, lParam
        End If
        ISubclass_WindowProc = 0
        
    Case WM_DRAWCLIPBOARD
        ' the content of the clipboard has changed.
        ' We raise a ClipboardChanged message and pass the message on:
        RaiseEvent ClipboardChanged
        SendMessageByLong m_hWndNextViewer, iMsg, wParam, lParam
        ISubclass_WindowProc = 0
        
    End Select
End Function

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = m_emr
End Property
Private Property Let ISubclass_MsgResponse(ByVal emrA As EMsgResponse)
    m_emr = emrA
End Property

