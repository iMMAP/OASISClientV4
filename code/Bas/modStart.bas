Attribute VB_Name = "modStart"
Option Explicit

Private Declare Function CoLockObjectExternal _
                Lib "ole32" (ByVal pUnk As IUnknown, _
                             ByVal fLock As Long, _
                             ByVal fLastUnlockReleases As Long) As Long

Private Declare Function SetTimer _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal nIDEvent As Long, _
                              ByVal uElapse As Long, _
                              ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal nIDEvent As Long) As Long

Const MaxLogSize = 2000000

Private m_colRunnables As Collection
Private m_lTimerID As Long

Public Sub WriteLog(sLogEntry As String, Optional bError As Boolean = False)

    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim sLogFile As String, sLogPath As String, iLogSize As Long
    Dim fso, f
    Dim sOldString As String
   
    On Error GoTo ErrHandler

    'Set the path and filename of the log
    sLogPath = App.Path & "\log\" & App.EXEName & " (" & Format(Now(), "yyyy-mm-dd") & IIf(bError, ") ERROR", ")")
    sLogFile = sLogPath & ".log"
   
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(sLogFile, ForAppending, True)
   
    'Get the size of the log to check if it's getting unwieldly
    iLogSize = GetLogSize(sLogFile)

    If iLogSize > MaxLogSize Then
   
        'If too big, back it up to to retain some sort of history
        fso.CopyFile sLogFile, (sLogPath & " OLD.log"), True
        Set f = Nothing
        fso.DeleteFile sLogFile
        'And start with a clean log-file
        Set f = fso.OpenTextFile(sLogFile, ForAppending, True)
        
    End If
    
    'Append the log-entry to the file together with time and date
    Dim oZLIB As New cZLIB
    sOldString = sLogEntry
    oZLIB.CompressString sLogEntry, Z_ZLIB
    f.WriteLine sLogEntry
    Set oZLIB = Nothing
    sLogEntry = sOldString
    
ErrHandler:
    Exit Sub
End Sub

Private Function GetLogSize(filespec As String) As Long
'Returns the size of a file in bytes. If the file does not
'exist, it returns -1.

   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   If (fso.FileExists(filespec)) Then
        Set f = fso.GetFile(filespec)
        GetLogSize = f.Size
   Else
        GetLogSize = -1
   End If
End Function

Public Function CheckInternConnection(ByRef sConnectionType As String) As Boolean
    Dim Flags As Long
    Dim result As Boolean

    sConnectionType = ""

    result = InternetGetConnectedState(Flags, 0)

    If result Then
        CheckInternConnection = True
    Else
        CheckInternConnection = False
    End If
     
    If Flags And INTERNET_CONNECTION_MODEM Then sConnectionType = "Connection Via Modem"
    If Flags And INTERNET_CONNECTION_LAN Then sConnectionType = "Connection Via LAN"
    If Flags And INTERNET_CONNECTION_PROXY Then sConnectionType = "Connection uses a Proxy"
    If Flags And INTERNET_CONNECTION_MODEM_BUSY Then sConnectionType = "Connection Via Modem but modem is busy"
    If Flags And INTERNET_CONNECTION_CONFIGURED Then sConnectionType = "Local system has a valid connection to the Internet, but it may or may not be currently connected."
    If Flags And INTERNET_CONNECTION_OFFLINE Then sConnectionType = "Local system is in offline mode."
    
End Function


Private Sub TimerProc(ByVal lHwnd As Long, _
                      ByVal lMsg As Long, _
                      ByVal lTimerID As Long, _
                      ByVal lTime As Long)
    Dim this As Runnable
    ' Go through the collection
    ' Runnable_Start method for each item in it
   
    With m_colRunnables

        Do While .Count > 0
            Set this = .Item(1)
            .Remove 1
            this.Start
            CoLockObjectExternal this, 0, 1
        Loop

    End With
   
    KillTimer 0, lTimerID
    m_lTimerID = 0
End Sub

Public Sub Start(this As Runnable)
    
    CoLockObjectExternal this, 1, 1

    ' Add this to collection
    If m_colRunnables Is Nothing Then
        Set m_colRunnables = New Collection
    End If

    m_colRunnables.Add this

    If Not m_lTimerID Then
        m_lTimerID = SetTimer(0, 0, 1, AddressOf TimerProc)
    End If

End Sub

