Attribute VB_Name = "DirectoryFunctions"
Option Explicit

Private Declare Function PathMatchSpec _
                Lib "shlwapi" _
                Alias "PathMatchSpecW" (ByVal pszFileParam As Long, _
                                        ByVal pszSpec As Long) As Long
   
Private Declare Function GetLogicalDriveStrings _
                Lib "kernel32" _
                Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                 ByVal lpBuffer As String) As Long
    
Private Declare Function FindExecutable _
                Lib "shell32.dll" _
                Alias "FindExecutableA" (ByVal lpFile As String, _
                                         ByVal lpDirectory As String, _
                                         ByVal lpResult As String) As Long

Function ScanTree(ByVal Path As String, Exclude As Boolean, ParamArray Mask()) As Variant

    On Error GoTo ERRORHANDLER
    
    Dim aCurrentFolderContents() As String
    Dim aFiles() As String
    Dim strCurrentFile As String
    Dim intFileCount As Long
    Dim intCurFile As Long
    Dim intParmCount As Integer
    Dim i As Integer
    Dim aMasks As Variant
    Dim bMatchesMask As Boolean
    
    'make sure the path ends with a "\"
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    
    intParmCount = UBound(Mask)

    If intParmCount <> -1 Then
        If intParmCount = 0 Then
            If IsArray(Mask(0)) Then
                aMasks = Mask(0)
            Else
                aMasks = Array()
                aAdd aMasks, Mask(0)
            End If

        Else
            aMasks = Array()

            For i = 0 To intParmCount
                aAdd aMasks, Mask(i)
            Next i

        End If
    End If
        
    'get all the subfolders
    strCurrentFile = GetFileQ(Path & "*.", vbDirectory)

    Do While strCurrentFile <> ""
        aAdd aCurrentFolderContents, strCurrentFile
        strCurrentFile = GetFileQ

        DoEvents
    Loop
    
    If Exclude Then
        strCurrentFile = GetFileQ(Path & "*.*", vbArchive + vbHidden + vbNormal + vbReadOnly)

        Do While strCurrentFile <> ""
            aAdd aCurrentFolderContents, strCurrentFile
            strCurrentFile = GetFileQ

            DoEvents
        Loop

    Else
        intParmCount = aLen(aMasks) - 1

        For i = 0 To intParmCount
            strCurrentFile = GetFileQ(Path & aMasks(i), vbArchive + vbHidden + vbNormal + vbReadOnly)

            Do While strCurrentFile <> ""
                aAdd aCurrentFolderContents, strCurrentFile
                strCurrentFile = GetFileQ

                DoEvents
            Loop

        Next i

    End If
    
    'adjust file count for zero-based array
    intFileCount = aLen(aCurrentFolderContents) - 1
    
    For intCurFile = 0 To intFileCount

        If UCase(aCurrentFolderContents(intCurFile)) <> "PAGEFILE.SYS" Then
            If GetAttr(Path & aCurrentFolderContents(intCurFile)) And vbDirectory Then
                If aCurrentFolderContents(intCurFile) <> "." And aCurrentFolderContents(intCurFile) <> ".." Then
                    'recurse subfolders
                    aConcat aFiles, ScanTree(Path & aCurrentFolderContents(intCurFile), Exclude, aMasks)
                End If

            Else
                
                bMatchesMask = True

                If Exclude Then

                    For i = 0 To intParmCount

                        If Not MatchesMask(Path & aCurrentFolderContents(intCurFile), CStr(aMasks(i)), Exclude) Then
                            bMatchesMask = False
                            Exit For
                        End If

                    Next i

                End If
            
                If bMatchesMask Then
                    aAdd aFiles, Path & aCurrentFolderContents(intCurFile)
                End If
                
            End If
        End If

        DoEvents
    Next
    
    ScanTree = aFiles
    
    Exit Function
    
ERRORHANDLER:
   
    Err.Raise vbObjectError + 1234, "ScanTree"
   
End Function

Private Function MatchesMask(sFile As String, _
                             sSpec As String, _
                             Optional Exclude As Boolean) As Boolean

    MatchesMask = CBool(PathMatchSpec(StrPtr(sFile), StrPtr(sSpec))) = Not Exclude
   
End Function

Function DriveExists(ByVal sDrive As String) As Boolean

    Dim buffer As String
    buffer = Space(64)

    If Len(sDrive) = 0 Then Exit Function
    GetLogicalDriveStrings Len(buffer), buffer
    DriveExists = InStr(1, buffer, Left$(sDrive, 1), vbTextCompare)
    
End Function

Function IsDriveReady(sDrive As String) As Boolean

    Dim fso As New FileSystemObject
    IsDriveReady = fso.GetDrive(sDrive).IsReady
    Set fso = Nothing
   
End Function

Function GetFileQ(Optional ByVal strPath, Optional lngType As FileAttribute) As String

    Dim strDrive As String
    
    If Not IsMissing(strPath) Then
    
        If InStr(strPath, ":") = 2 Then
            strDrive = Left(strPath, 1)
        Else
            strDrive = Left(App.Path, 1)
        End If
        
        If Not DriveExists(strDrive) Then Exit Function
        
        If Not IsDriveReady(strDrive) Then Exit Function
        
        GetFileQ = Dir(strPath, lngType)
        
    Else
    
        GetFileQ = Dir()
        
    End If
    
End Function

Function GetFileList(strSourceDir As String, bIncludeSubs As Boolean, Optional strMask As String) As Variant

    Dim strFileName     As String
    Dim aFileList       As Variant
    
    Dim t1 As Single
    Dim totaltime As Single

    If Right(strSourceDir, 1) <> "\" Then strSourceDir = strSourceDir & "\"
    If strMask = "" Then strMask = "*.*"

   ' t1 = Timer

    If bIncludeSubs Then
        aFileList = ScanTree(strSourceDir, False, strMask)
    Else
        aFileList = Array()
        strFileName = GetFileQ(strSourceDir & strMask, vbNormal + vbReadOnly + vbHidden)

        Do While strFileName <> ""
            aAdd aFileList, strSourceDir & strFileName
            strFileName = GetFileQ

            DoEvents
        Loop

    End If
    
    GetFileList = aFileList
    
    'totaltime = Timer - t1

End Function

Public Function GetAssociatedExecutable(ByVal Filename As String) As String

    Const MAX_PATH As Long = 260
    Dim Path As String

    Path = String$(MAX_PATH, 0)

    Call FindExecutable(Filename, vbNullString, Path)
        
    GetAssociatedExecutable = Left$(Path, InStr(Path, vbNullChar) - 1)

End Function

