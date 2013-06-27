VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Enum EMemoryFlags
    GMEM_DDESHARE = &H2000
    GMEM_DISCARDABLE = &H100
    GMEM_DISCARDED = &H4000
    GMEM_INVALID_HANDLE = &H8000
    GMEM_FIXED = &H0
    GMEM_LOCKCOUNT = &HFF
    GMEM_MODIFY = &H80
    GMEM_MOVEABLE = &H2
    GMEM_NODISCARD = &H20
    GMEM_NOCOMPACT = &H10
    GMEM_NOT_BANKED = &H1000
    GMEM_LOWER = GMEM_NOT_BANKED
    GMEM_NOTIFY = &H4000
    GMEM_SHARE = &H2000
    GMEM_VALID_FLAGS = &H7F72
    GMEM_ZEROINIT = &H40
    GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
End Enum

Private m_hMem As Long
Private m_lPtr As Long

Public Property Get Handle() As Long
        '<EhHeader>
        On Error GoTo Handle_Err
        '</EhHeader>
100     Handle = m_hMem
        '<EhFooter>
        Exit Property

Handle_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cMemory.Handle", _
                  "cMemory component failure"
        '</EhFooter>
End Property
Public Property Let Handle(ByVal hMem As Long)
        '<EhHeader>
        On Error GoTo Handle_Err
        '</EhHeader>
100     If (m_hMem <> 0) Then
102         FreeMemory
        End If
104     m_hMem = hMem
        '<EhFooter>
        Exit Property

Handle_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cMemory.Handle", _
                  "cMemory component failure"
        '</EhFooter>
End Property
Public Property Get Pointer() As Long
        '<EhHeader>
        On Error GoTo Pointer_Err
        '</EhHeader>
100     If (m_hMem <> 0) Then
102         If (m_lPtr = 0) Then
104             LockMemory
            End If
106         Pointer = m_lPtr
        End If
        '<EhFooter>
        Exit Property

Pointer_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cMemory.Pointer", _
                  "cMemory component failure"
        '</EhFooter>
End Property
Public Property Get Size() As Long
        '<EhHeader>
        On Error GoTo Size_Err
        '</EhHeader>
100     If (m_hMem <> 0) Then
102         Size = GlobalSize(m_hMem)
        End If
        '<EhFooter>
        Exit Property

Size_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cMemory.Size", _
                  "cMemory component failure"
        '</EhFooter>
End Property
Public Function AllocateMemory( _
        ByVal lSize As Long, _
        Optional ByVal dwFlags As Long = GPTR _
    ) As Boolean
        '<EhHeader>
        On Error GoTo AllocateMemory_Err
        '</EhHeader>
100     FreeMemory
102     m_hMem = GlobalAlloc(dwFlags, lSize)
104     If (m_hMem <> 0) Then
            ' Success
106         AllocateMemory = True
        Else
            ' Failed
        End If
        '<EhFooter>
        Exit Function

AllocateMemory_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cMemory.AllocateMemory", _
                  "cMemory component failure"
        '</EhFooter>
End Function
Public Function LockMemory() As Boolean
        '<EhHeader>
        On Error GoTo LockMemory_Err
        '</EhHeader>
100     If (m_hMem <> 0) Then
102         If (m_lPtr = 0) Then
104             m_lPtr = GlobalLock(m_hMem)
106             If (m_lPtr <> 0) Then
                    ' Success
108                 LockMemory = True
                Else
                    ' Failed
                End If
            End If
        End If
        '<EhFooter>
        Exit Function

LockMemory_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cMemory.LockMemory", _
                  "cMemory component failure"
        '</EhFooter>
End Function
Public Sub UnlockMemory()
        '<EhHeader>
        On Error GoTo UnlockMemory_Err
        '</EhHeader>
100     If (m_hMem <> 0) Then
102         If (m_lPtr <> 0) Then
104             GlobalUnlock m_hMem
106             m_lPtr = 0
            End If
        End If
        '<EhFooter>
        Exit Sub

UnlockMemory_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cMemory.UnlockMemory", _
                  "cMemory component failure"
        '</EhFooter>
End Sub
Public Sub FreeMemory()
        '<EhHeader>
        On Error GoTo FreeMemory_Err
        '</EhHeader>
100     If (m_hMem <> 0) Then
102         UnlockMemory
104         GlobalFree m_hMem
        End If
106     m_hMem = 0
        '<EhFooter>
        Exit Sub

FreeMemory_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cMemory.FreeMemory", _
                  "cMemory component failure"
        '</EhFooter>
End Sub
Public Sub ReleaseDontFreeMemory()
        ' For GMEM_DDESHARE operations...
        '<EhHeader>
        On Error GoTo ReleaseDontFreeMemory_Err
        '</EhHeader>
100     UnlockMemory
102     m_hMem = 0
        '<EhFooter>
        Exit Sub

ReleaseDontFreeMemory_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cMemory.ReleaseDontFreeMemory", _
                  "cMemory component failure"
        '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    FreeMemory
End Sub


