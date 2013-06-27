Attribute VB_Name = "FolderMonitorGlobals"
Option Explicit

Public Const FILE_NOTIFY_CHANGE_FILE_NAME As Long = &H1
Public Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2
Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4
Public Const FILE_NOTIFY_CHANGE_SIZE As Long = &H8
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE As Long = &H10
Public Const FILE_NOTIFY_CHANGE_LAST_ACCESS As Long = &H20
Public Const FILE_NOTIFY_CHANGE_CREATION As Long = &H40
Public Const FILE_NOTIFY_CHANGE_SECURITY As Long = &H100
Public Const FILE_NOTIFY_FLAGS = FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                          FILE_NOTIFY_CHANGE_FILE_NAME Or _
                          FILE_NOTIFY_CHANGE_LAST_WRITE

Public Const INVALID_HANDLE_VALUE = -1
Public Const SYNCHRONIZE = &H100000
Public Const WM_CLOSE = &H10

Public Const WAIT_OBJECT = &H0
Public Const WAIT_TIMEOUT = &H102
Public Const WAIT_TIME = 100
Public Const WAIT_FAILED = &HFFFFFFFF
Public Const INFINITE = -1&
Public Const MAXIMUM_WAIT_OBJECTS = &H40

Public Declare Function FindFirstChangeNotification Lib "kernel32" _
    Alias "FindFirstChangeNotificationA" _
   (ByVal lpPathName As String, _
    ByVal bWatchSubtree As Long, _
    ByVal dwNotifyFilter As Long) As Long

Public Declare Function FindNextChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

Public Declare Function FindCloseChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

Public Declare Function WaitForMultipleObjects Lib "kernel32" ( _
    ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll _
    As Long, ByVal dwMilliseconds As Long) As Long

