VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThreads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 0

Private Type udtThread
    Handle As Long
    Enabled As Boolean
End Type

Private uThread As udtThread

Private Const CREATE_SUSPENDED As Long = &H4
Private Const THREAD_BASE_PRIORITY_IDLE  As Long = -15
Private Const THREAD_BASE_PRIORITY_LOWRT  As Long = 15
Private Const THREAD_BASE_PRIORITY_MAX As Long = 2
Private Const THREAD_BASE_PRIORITY_MIN As Long = -2
Private Const THREAD_PRIORITY_HIGHEST  As Long = THREAD_BASE_PRIORITY_MAX
Private Const THREAD_PRIORITY_LOWEST  As Long = THREAD_BASE_PRIORITY_MIN
Private Const THREAD_PRIORITY_ABOVE_NORMAL As Long = (THREAD_PRIORITY_HIGHEST - 1)
Private Const THREAD_PRIORITY_BELOW_NORMAL  As Long = (THREAD_PRIORITY_LOWEST + 1)
Private Const THREAD_PRIORITY_IDLE  As Long = THREAD_BASE_PRIORITY_IDLE
Private Const THREAD_PRIORITY_NORMAL As Long = 0
Private Const THREAD_PRIORITY_TIME_CRITICAL As Long = THREAD_BASE_PRIORITY_LOWRT

Private Declare Function CreateThread _
                Lib "Kernel32" (ByVal lpThreadAttributes As Any, _
                                ByVal dwStackSize As Long, _
                                ByVal lpStartAddress As Long, _
                                lpParameter As Any, _
                                ByVal dwCreationFlags As Long, _
                                lpThreadId As Long) As Long
Private Declare Function ResumeThread _
                Lib "Kernel32" (ByVal hThread As Long) As Long
Private Declare Function SetThreadPriority _
                Lib "Kernel32" (ByVal hThread As Long, _
                                ByVal nPriority As Long) As Long
Private Declare Function GetThreadPriority _
                Lib "Kernel32" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread _
                Lib "Kernel32" (ByVal hThread As Long) As Long
Private Declare Function TerminateThread _
                Lib "Kernel32" (ByVal hThread As Long, _
                                ByVal dwExitCode As Long) As Long



Public Sub Initialize(ByVal lpfnBasFunc As Long)

    Dim lStackSize As Long, lCreationFlags As Long, lpThreadId As Long, lNull As Long
    
    On Error Resume Next
    lNull = 0
    lStackSize = 0
    lCreationFlags = CREATE_SUSPENDED
    uThread.Handle = CreateThread(lNull, lStackSize, lpfnBasFunc, lNull, lCreationFlags, lpThreadId)

    If uThread.Handle = lNull Then MsgBox "Create thread failed!"
    
End Sub

Public Property Get Enabled() As Boolean
    
    On Error Resume Next
    Enabled = uThread.Enabled

End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    
    On Error Resume Next

    If vNewValue And (Not uThread.Enabled) Then
        ResumeThread uThread.Handle
        uThread.Enabled = True
    ElseIf uThread.Enabled Then
        SuspendThread uThread.Handle
        uThread.Enabled = False
    End If

End Property

Public Property Get Priority() As Long

    On Error Resume Next
    Priority = GetThreadPriority(uThread.Handle)

End Property

Public Property Let Priority(ByVal vNewValue As Long)

    On Error Resume Next

    If vNewValue = -2 Then
        Call SetThreadPriority(uThread.Handle, THREAD_PRIORITY_LOWEST)
    ElseIf vNewValue = -1 Then
        Call SetThreadPriority(uThread.Handle, THREAD_PRIORITY_BELOW_NORMAL)
    ElseIf vNewValue = 0 Then
        Call SetThreadPriority(uThread.Handle, THREAD_PRIORITY_NORMAL)
    ElseIf vNewValue = 1 Then
        Call SetThreadPriority(uThread.Handle, THREAD_PRIORITY_ABOVE_NORMAL)
    ElseIf vNewValue = 2 Then
        Call SetThreadPriority(uThread.Handle, THREAD_PRIORITY_HIGHEST)
    End If

End Property

Private Sub Class_Terminate()

    On Error Resume Next
    Call TerminateThread(uThread.Handle, 0)

End Sub
