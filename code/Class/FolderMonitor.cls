VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FldMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Event ChangeOccurred()

Dim bStopped            As Boolean
Dim bIncludeSubs        As Boolean
Dim strMask             As String
Dim aFolders()          As String
Dim aHandles()          As Long
Dim aChangedFolders     As Variant
Dim lngMilliSeconds     As Long
Dim bChanged            As Boolean
Dim strLastNotification As String
Dim bMonitorAttributes  As Boolean
Dim bMonitorDeleted     As Boolean
Dim FILE_NOTIFY_FLAGS   As Long

Dim objFolderHash       As Dictionary

Private Function WaitForChange() As Long

    Dim lngResult       As Long
    
    lngResult = WAIT_TIMEOUT
    Do While lngResult = WAIT_TIMEOUT And Not bStopped
        lngResult = WaitForMultipleObjects(aLen(aHandles), aHandles(0), False, lngMilliSeconds)
        DoEvents
    Loop
    
    WaitForChange = lngResult
    
End Function

Sub StopMonitoring()
    
    Dim i           As Integer
    Dim intHandles  As Integer
    
    intHandles = aLen(aHandles) - 1
    For i = 0 To intHandles
        Call FindCloseChangeNotification(aHandles(i))
    Next i
    
    bStopped = True
    
    If Not objFolderHash Is Nothing Then objFolderHash.RemoveAll
    Set objFolderHash = Nothing
    
    DoEvents
    
End Sub

Sub StartMonitoring()

    Dim lngStatus           As Long
    Dim aCurrentFiles       As Variant
    Dim lngFileCount        As Long
    Dim lngFolderHandle     As Long
    Dim l                   As Long
    Dim strFileInfo         As String
    Dim i                   As Integer
    Dim intFolders          As Integer
        
    On Error Resume Next
    
    Set objFolderHash = New Dictionary
        
    intFolders = aLen(aFolders) - 1
    For i = 0 To intFolders
    
        If GetFileQ(aFolders(i), vbDirectory) = "" Then
            Err.Raise Err.Number, "FileNotification.StartMonitoring()", "Invalid path specified for monitoring " & aFolders(i)
        End If
        
        aCurrentFiles = GetFileList(aFolders(i), bIncludeSubs, strMask)
        lngFileCount = aLen(aCurrentFiles) - 1
        
        For l = 0 To lngFileCount
            strFileInfo = CStr(aCurrentFiles(l)) & Chr(255) & CStr(FileDateTime(aCurrentFiles(l))) & IIf(bMonitorAttributes, CStr(GetAttr(aCurrentFiles(l))), "")
            objFolderHash.Add strFileInfo, ""
            DoEvents
        Next l
        
        bStopped = False
        
        lngFolderHandle = FindFirstChangeNotification(aFolders(i), bIncludeSubs, FILE_NOTIFY_FLAGS)
        If lngFolderHandle = INVALID_HANDLE_VALUE Then
            Err.Raise Err.Number, "FileNotification.StartMonitoring()", "Failed to initialize folder for monitoring"
        End If
        
        aAdd aHandles, lngFolderHandle
        
    Next
        
    lngStatus = WaitForChange
        
    Select Case lngStatus
    
        Case WAIT_FAILED
            If Not bStopped Then
                Err.Raise Err.Number, "FolderNotification.StartMonitoring()"
            End If
            StopMonitoring

        Case Else
        
            DoEvents
            If aScan(aChangedFolders, lngStatus) = -1 Then
                aAdd aChangedFolders, lngStatus
            End If
            bChanged = True
            RaiseEvent ChangeOccurred
                    
            Do While lngStatus <> WAIT_TIMEOUT And Not bStopped
            
                For i = 0 To intFolders
                    Call FindNextChangeNotification(aHandles(i))
                Next i
                lngStatus = WaitForChange()
        
                Select Case lngStatus
                    
                    Case WAIT_FAILED
                        If Not bStopped Then
                            Err.Raise Err.Number, "FolderNotification.StartMonitoring()", "Wait failed"
                        End If
                        StopMonitoring
                        
                    Case Else
                        DoEvents
                        If aScan(aChangedFolders, lngStatus) = -1 Then
                            aAdd aChangedFolders, lngStatus
                        End If
                        bChanged = True
                        RaiseEvent ChangeOccurred
             
                End Select
                
            Loop
        
    End Select
    
End Sub

Public Property Get isMonitoring() As Boolean

    isMonitoring = Not bStopped
    
End Property

Public Property Let isMonitoring(bRunning As Boolean)

     bStopped = Not bRunning
    
End Property

Public Function ChangedList() As Variant

    Dim aCurrentFileList    As Variant
    Dim lngCurrentFileCount As Long
    Dim l                   As Long
    Dim strFileInfo         As String
    Dim intFolders          As Integer
    Dim i                   As Integer
    Dim lngFileCount        As Long
    Dim aFiles              As Variant
        
    Dim objCurrentFiles     As Dictionary
    
    ChangedList = Array()
    
    If Not isMonitoring Or Not Changed Then Exit Function
    bChanged = False
    
    intFolders = aLen(aChangedFolders) - 1
    For i = 0 To intFolders
        aCurrentFileList = GetFileList(aFolders(aChangedFolders(i)), bIncludeSubs, strMask)
        lngCurrentFileCount = aLen(aCurrentFileList) - 1
        For l = 0 To lngCurrentFileCount
            strFileInfo = CStr(aCurrentFileList(l)) & Chr(255) 'add a delimiter so that we can retrieve the filename later
            
            If GetFileQ(aCurrentFileList(l)) <> "" Then
                strFileInfo = strFileInfo & CStr(FileDateTime(aCurrentFileList(l)))
            End If
            strFileInfo = strFileInfo & IIf(bMonitorAttributes, CStr(GetAttr(aCurrentFileList(l))), "")
            
            If Not objFolderHash.Exists(strFileInfo) Then
                aAdd ChangedList, CStr(aCurrentFileList(l))
                objFolderHash.Add strFileInfo, ""
            End If
            DoEvents
        Next
        
        If bMonitorDeleted Then
            Set objCurrentFiles = New Dictionary
            For l = 0 To lngCurrentFileCount
                objCurrentFiles.Add aCurrentFileList(l), ""
            Next l
            
            lngFileCount = objFolderHash.Count - 1
            aFiles = objFolderHash.Keys
            For l = 0 To lngFileCount
                strFileInfo = Left(aFiles(l), InStr(aFiles(l), Chr(255)) - 1)
                If Not objCurrentFiles.Exists(strFileInfo) Then
                    aAdd ChangedList, strFileInfo
                    objFolderHash.Remove aFiles(l)
                End If
            Next l
        End If
        
    Next i
    
    aChangedFolders = Array()
    
    Set objCurrentFiles = Nothing

End Function

Public Function AddFolder(ParamArray strPaths())
    
    
    Dim intParmCount    As Integer
    Dim i               As Integer
    
    If isMonitoring Then
        Err.Raise vbObjectError + 7854, "FolderMonitor.Folder", "Can not add folder to monitor while monitoring in progess"
    End If
    
    'PADDY NOTE!!!! the limit on the number of threads capable of being monitored simultaneously
    'by the Windows' API function WaitForMultipleObjects is 63
    If aLen(aFolders) = MAXIMUM_WAIT_OBJECTS - 1 Then
        Err.Raise vbObjectError + 7855, "FolderMonitor.Folder", "Only 63 folders can be monitored simultaneously"
    End If
    
    intParmCount = UBound(strPaths)
    If intParmCount <> -1 Then
        For i = 0 To intParmCount
            If GetFileQ(strPaths(i), vbDirectory) = "" Then
                Err.Raise Err.Number, "FolderMonitor.Folder", "Invalid path specified for monitoring " & strPaths(i)
            End If
            aAdd aFolders, Trim(strPaths(i))
        Next i
    End If
        
End Function

Public Property Get Folders() As Variant

    Folders = aFolders
    
End Property

Public Property Let IncludeSubFolders(bYesOrNo As Boolean)

    bIncludeSubs = bYesOrNo
    
End Property

Public Property Get IncludeSubFolders() As Boolean

    IncludeSubFolders = bIncludeSubs
    
End Property

Public Property Let Mask(strFileMask As String)

    strMask = strFileMask
    
End Property

Public Property Get Mask() As String

    Mask = strMask
    
End Property

Public Property Let WaitTime(msToWait As Long)

    lngMilliSeconds = msToWait

End Property

Public Property Get WaitTime() As Long

    WaitTime = lngMilliSeconds

End Property
Public Property Get Changed() As Boolean

    Changed = bChanged

End Property

Public Property Let Changed(bTrue As Boolean)

    bChanged = bTrue
    
End Property

Private Sub Class_Initialize()

    bStopped = True
    lngMilliSeconds = WAIT_TIME
    aChangedFolders = Array()
    
    FILE_NOTIFY_FLAGS = FILE_NOTIFY_CHANGE_FILE_NAME Or FILE_NOTIFY_CHANGE_LAST_WRITE
    If Attributes Then
        FILE_NOTIFY_FLAGS = FILE_NOTIFY_FLAGS Or FILE_NOTIFY_CHANGE_ATTRIBUTES
    End If
    
End Sub

Public Property Let Attributes(bOn As Boolean)

    bMonitorAttributes = bOn
    
    If bOn Then
        FILE_NOTIFY_FLAGS = FILE_NOTIFY_FLAGS Or FILE_NOTIFY_CHANGE_ATTRIBUTES
    End If

End Property

Public Property Get Attributes() As Boolean

    Attributes = bMonitorAttributes
    
End Property

Public Property Let Deleted(bOn As Boolean)

    bMonitorDeleted = bOn
    
End Property

Public Property Get Deleted() As Boolean

    Deleted = bMonitorDeleted
    
End Property

Private Sub Class_Terminate()

    Set objFolderHash = Nothing

End Sub
