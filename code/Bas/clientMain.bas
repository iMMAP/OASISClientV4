Attribute VB_Name = "clientMain"

Option Explicit
Public Const LOCALE_SSHORTDATE = &H1F

Public Declare Function GetSystemDefaultLCID _
               Lib "kernel32" () As Long

Public Declare Function SetLocaleInfo _
               Lib "kernel32" _
               Alias "SetLocaleInfoA" (ByVal Locale As Long, _
                                       ByVal LCType As Long, _
                                       ByVal lpLCData As String) As Boolean

Private Declare Function FindFirstFile _
                Lib "kernel32" _
                Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                        lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile _
                Lib "kernel32" _
                Alias "FindNextFileA" (ByVal hFindFile As Long, _
                                       lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes _
                Lib "kernel32" _
                Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose _
                Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Dim m_frmMain As New frmMain

Private Const DRIVE_UNKNOWN = 0
Private Const DRIVE_ABSENT = 1
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6
' returns errors for UNC Path
Private Const ERROR_BAD_DEVICE = 1200&
Private Const ERROR_CONNECTION_UNAVAIL = 1201&
Private Const ERROR_EXTENDED_ERROR = 1208&
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_NOT_SUPPORTED = 50&
Private Const ERROR_NO_NET_OR_BAD_PATH = 1203&
Private Const ERROR_NO_NETWORK = 1222&
Private Const ERROR_NOT_CONNECTED = 2250&
Private Const NO_ERROR = 0

Private Declare Function SetForegroundWindow _
                Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function FindWindow _
                Lib "user32" _
                Alias "FindWindowA" (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long

Private Declare Function WNetGetConnection _
                Lib "mpr.dll" _
                Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, _
                                            ByVal lpszRemoteName As String, _
                                            cbRemoteName As Long) As Long
Private Declare Function GetLogicalDriveStrings _
                Lib "kernel32" _
                Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                 ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType _
                Lib "kernel32" _
                Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation _
                Lib "kernel32" _
                Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                               ByVal lpVolumeNameBuffer As String, _
                                               ByVal nVolumeNameSize As Long, _
                                               lpVolumeSerialNumber As Long, _
                                               lpMaximumComponentLength As Long, _
                                               lpFileSystemFlags As Long, _
                                               ByVal lpFileSystemNameBuffer As String, _
                                               ByVal nFileSystemNameSize As Long) As Long
'Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Const MAX_FILENAME_LEN = 256

Const MOVEFILE_REPLACE_EXISTING = &H1
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_BEGIN = 0
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const CREATE_NEW = 1
Const OPEN_EXISTING = 3
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Private Declare Function SetVolumeLabel _
                Lib "kernel32" _
                Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, _
                                         ByVal lpVolumeName As String) As Long
Private Declare Function WriteFile _
                Lib "kernel32" (ByVal hFile As Long, _
                                lpBuffer As Any, _
                                ByVal nNumberOfBytesToWrite As Long, _
                                lpNumberOfBytesWritten As Long, _
                                ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile _
                Lib "kernel32" (ByVal hFile As Long, _
                                lpBuffer As Any, _
                                ByVal nNumberOfBytesToRead As Long, _
                                lpNumberOfBytesRead As Long, _
                                ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile _
                Lib "kernel32" _
                Alias "CreateFileA" (ByVal lpFileName As String, _
                                     ByVal dwDesiredAccess As Long, _
                                     ByVal dwShareMode As Long, _
                                     ByVal lpSecurityAttributes As Any, _
                                     ByVal dwCreationDisposition As Long, _
                                     ByVal dwFlagsAndAttributes As Long, _
                                     ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle _
                Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer _
                Lib "kernel32" (ByVal hFile As Long, _
                                ByVal lDistanceToMove As Long, _
                                lpDistanceToMoveHigh As Long, _
                                ByVal dwMoveMethod As Long) As Long
Private Declare Function SetFileAttributes _
                Lib "kernel32" _
                Alias "SetFileAttributesA" (ByVal lpFileName As String, _
                                            ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileSize _
                Lib "kernel32" (ByVal hFile As Long, _
                                lpFileSizeHigh As Long) As Long
Private Declare Function GetTempFileName _
                Lib "kernel32" _
                Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                          ByVal lpPrefixString As String, _
                                          ByVal wUnique As Long, _
                                          ByVal lpTempFileName As String) As Long
Private Declare Function MoveFileEx _
                Lib "kernel32" _
                Alias "MoveFileExA" (ByVal lpExistingFileName As String, _
                                     ByVal lpNewFileName As String, _
                                     ByVal dwFlags As Long) As Long
Private Declare Function DeleteFile _
                Lib "kernel32" _
                Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Declare Function LockWindowUpdate _
               Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function timeGetTime _
               Lib "winmm.dll" () As Long
Public ClipBoard_Ext As cCustomClipboard

Private Sub SetVolLabel()

    Dim sSave As String, hOrgFile As Long, hNewFile As Long, bBytes() As Byte
    Dim sTemp As String, nSize As Long, Ret As Long
    'Ask for a new volume label
    sSave = InputBox("Please enter a new volume label for drive C:\" + vbCrLf + " (if you don't want to change it, leave the textbox blank)")

    If sSave <> "" Then
        SetVolumeLabel "C:\", sSave
    End If

    'Create a buffer
    sTemp = String(260, 0)
    'Get a temporary filename
    GetTempFileName "C:\", "KPD", 0, sTemp
    'Remove all the unnecessary chr$(0)'s
    sTemp = left$(sTemp, InStr(1, sTemp, Chr$(0)) - 1)
    'Set the file attributes
    SetFileAttributes sTemp, FILE_ATTRIBUTE_TEMPORARY
    'Open the files
    hNewFile = CreateFile(sTemp, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    hOrgFile = CreateFile("c:\config.sys", GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    'Get the file size
    nSize = GetFileSize(hOrgFile, 0)
    'Set the file pointer
    SetFilePointer hOrgFile, Int(nSize / 2), 0, FILE_BEGIN
    'Create an array of bytes
    ReDim bBytes(1 To nSize - Int(nSize / 2)) As Byte
    'Read from the file
    ReadFile hOrgFile, bBytes(1), UBound(bBytes), Ret, ByVal 0&

    'Check for errors
    If Ret <> UBound(bBytes) Then MsgBox "Error reading file ..."

    'Write to the file
    WriteFile hNewFile, bBytes(1), UBound(bBytes), Ret, ByVal 0&

    'Check for errors
    If Ret <> UBound(bBytes) Then MsgBox "Error writing file ..."

    'Close the files
    CloseHandle hOrgFile
    CloseHandle hNewFile

    'Move the file
    MoveFileEx sTemp, "C:\KPDTEST.TST", MOVEFILE_REPLACE_EXISTING
    'Delete the file
    DeleteFile "C:\KPDTEST.TST"
    'Unload Me
End Sub

Public Function DriveSerial(ByVal sDrv As String) As Long
    Dim RetVal As Long
    Dim Str As String * MAX_FILENAME_LEN
    Dim str2 As String * MAX_FILENAME_LEN
    Dim a As Long
    Dim b As Long
    Call GetVolumeInformation(sDrv & ":\", Str, MAX_FILENAME_LEN, RetVal, a, b, str2, MAX_FILENAME_LEN)
    DriveSerial = RetVal
End Function

Private Sub FgetSerial()
    MsgBox "Serial of drive C is " & DriveSerial("C")
End Sub

Private Function fGetDrives() As String
    'Returns all mapped drives
    Dim lngRet As Long
    Dim strDrives As String * 255
    Dim lngTmp As Long
    lngTmp = Len(strDrives)
    lngRet = GetLogicalDriveStrings(lngTmp, strDrives)
    fGetDrives = left(strDrives, lngRet)
End Function

Private Function fGetUNCPath(strDriveLetter As String) As String
    On Local Error GoTo fGetUNCPath_Err

    Dim msg As String, lngReturn As Long
    Dim lpszLocalName As String
    Dim lpszRemoteName As String
    Dim cbRemoteName As Long
    lpszLocalName = strDriveLetter
    lpszRemoteName = String$(255, Chr$(32))
    cbRemoteName = Len(lpszRemoteName)
    lngReturn = WNetGetConnection(lpszLocalName, lpszRemoteName, cbRemoteName)

    Select Case lngReturn

        Case ERROR_BAD_DEVICE
            msg = "Error: Bad Device"

        Case ERROR_CONNECTION_UNAVAIL
            msg = "Error: Connection Un-Available"

        Case ERROR_EXTENDED_ERROR
            msg = "Error: Extended Error"

        Case ERROR_MORE_DATA
            msg = "Error: More Data"

        Case ERROR_NOT_SUPPORTED
            msg = "Error: Feature not Supported"

        Case ERROR_NO_NET_OR_BAD_PATH
            msg = "Error: No Network Available or Bad Path"

        Case ERROR_NO_NETWORK

            msg = "Error: No Network Available"

        Case ERROR_NOT_CONNECTED
            msg = "Error: Not Connected"

        Case NO_ERROR
            ' all is successful...
    End Select

    If Len(msg) Then
        MsgBox msg, vbInformation
    Else
        fGetUNCPath = left$(lpszRemoteName, cbRemoteName)
    End If

fGetUNCPath_End:
    Exit Function
fGetUNCPath_Err:
    MsgBox Err.Description, vbInformation
    Resume fGetUNCPath_End
End Function

Private Function fDriveType(strDriveName As String) As String
    Dim lngRet As Long
    Dim strDrive As String
    lngRet = GetDriveType(strDriveName)

    Select Case lngRet

        Case DRIVE_UNKNOWN 'The drive type cannot be determined.
            strDrive = "Unknown Drive Type"

        Case DRIVE_ABSENT 'The root directory does not exist.
            strDrive = "Drive does not exist"

        Case DRIVE_REMOVABLE 'The drive can be removed from the drive.
            strDrive = "Removable Media"

        Case DRIVE_FIXED 'The disk cannot be removed from the drive.
            strDrive = "Fixed Drive"

        Case DRIVE_REMOTE  'The drive is a remote (network) drive.
            strDrive = "Network Drive"

        Case DRIVE_CDROM 'The drive is a CD-ROM drive.
            strDrive = "CD Rom"

        Case DRIVE_RAMDISK 'The drive is a RAM disk.
            strDrive = "Ram Disk"
    End Select

    fDriveType = strDrive
End Function

Sub sListAllDrives()

    Dim strAllDrives As String
    Dim strTmp As String
    Dim Serial As Long, VName As String, FSName As String
    Dim mFile As File, mFileSysObj As New FileSystemObject, mTxtStream As TextStream
    Dim outputText As String
    Dim Txt As String
    
    If Not mFileSysObj.FileExists(g_sAppPath & "\data\user\Sessions\guid.dat") Then
        mFileSysObj.CreateTextFile g_sAppPath & "\data\user\Sessions\guid.dat", True
    End If
    
    Set mFile = mFileSysObj.GetFile(g_sAppPath & "\data\user\Sessions\guid.dat")
    Set mTxtStream = mFile.OpenAsTextStream(IOMode.ForReading)
     
    If Not mTxtStream.AtEndOfLine Then
        Txt = mTxtStream.ReadAll
    End If
    
    strTmp = left(CStr(g_sAppPath), 3)
    VolInfo strTmp, Serial, VName, FSName
    
    If Len(Txt) < 1 Then
        Set mTxtStream = mFile.OpenAsTextStream(IOMode.ForWriting)
        mTxtStream.WriteLine "serial=" & Serial & vbCrLf & "name=" & VName
        mTxtStream.Close
    Else
        Set mTxtStream = mFile.OpenAsTextStream(IOMode.ForReading)
        Txt = mTxtStream.ReadLine
        
        If Txt = ("serial=" & CStr(Serial)) Then
            Txt = mTxtStream.ReadLine

            If CStr(("name=" & VName)) = Txt Then
                DebugPrint "OK"
            Else
                MsgBox "It seems like the OASIS client has been copied. Or removed from it's original location. The hardware label has been changed since last time." & vbCrLf & "Please contatct your OASIS administrator to achieve a new version.", vbCritical, "OASIS Client"
                End
            End If

        Else
            MsgBox "It seems like the OASIS client has been copied. Or removed from it's original location. The Hardware ID has been changed." & vbCrLf & "Please contatct your OASIS administrator to achieve a new version.", vbCritical, "OASIS Client"
            End
        End If
        
        mTxtStream.Close
    End If
    
    'MsgBox strTmp & vbCrLf & Serial & vbCrLf & VName & vbCrLf & FSName
                              
    Exit Sub
                              
    strAllDrives = fGetDrives

    If strAllDrives <> "" Then

        Do
            strTmp = Mid$(strAllDrives, 1, InStr(strAllDrives, vbNullChar) - 1)
            strAllDrives = Mid$(strAllDrives, InStr(strAllDrives, vbNullChar) + 1)

            Select Case fDriveType(strTmp)

                Case "Removable Media":
                    DebugPrint "Removable drive :  " & strTmp
                    VolInfo strTmp, Serial, VName, FSName
                    MsgBox strTmp & vbCrLf & Serial & vbCrLf & VName & vbCrLf & FSName
                    
                Case "CD Rom":
                    DebugPrint "   CD Rom drive :  " & strTmp

                Case "Fixed Drive":
                    DebugPrint "    Local drive :  " & strTmp

                Case "Network Drive":
                    'DebugPrint "  Network drive :  " & strTmp
                    'DebugPrint "       UNC Path :  " & _
                    '            fGetUNCPath(Left$(strTmp, Len(strTmp) - 1))
            End Select

        Loop While strAllDrives <> ""

    End If

End Sub

'Private Sub Form_Load()
'    DebugPrint "All available drives: "
'    sListAllDrives
'End Sub

Private Sub VolInfo(sPath As String, _
                    Serial As Long, _
                    VName As String, _
                    FSName As String)
    'Create buffers
    VName = String$(255, Chr$(0))
    FSName = String$(255, Chr$(0))
    'Get the volume information
    GetVolumeInformation sPath, VName, 255, Serial, 0, 0, FSName, 255
    'Strip the extra chr$(0)'s
    VName = left$(VName, InStr(1, VName, Chr$(0)) - 1)
    FSName = left$(FSName, InStr(1, FSName, Chr$(0)) - 1)
    DebugPrint "The Volume name of " & sPath & " is '" + VName + "', the File system name of " & sPath & " is '" + FSName + "' and the serial number of " & sPath & " is '" + Trim(Str$(Serial)) + "'"
End Sub

Private Function IsFormLoaded(FormName As String, _
                              frm As Form) As Boolean
    'Returns True if form name is found
    Dim f As Form

    For Each f In Forms

        If f.Name = FormName Then
            Set frm = f
            IsFormLoaded = True
            Exit For
        End If

    Next f

End Function

Public Sub LoadLanguage(sfrmName As String, _
                        sLanguage As String, _
                        CN As ADODB.Connection, _
                        Optional bLoad As Boolean, _
                        Optional bShow As Boolean, _
                        Optional iModal As Integer, _
                        Optional lWHND As Long)
    Dim rslang As ADODB.Recordset: Set rslang = New ADODB.Recordset
    Dim frm As Form
    Dim ctr As Control
    
    On Error GoTo Hell
    
    If Not IsFormLoaded(sfrmName, frm) Then
            
        If Not bLoad Then Exit Sub
            
        Set frm = Forms.Add(sfrmName)
    End If
    
    rslang.Open "Select * FROM lang WHERE Container = '" & sfrmName & "'", CN, adOpenDynamic, adLockOptimistic 'm_Cnn
    
    If Not rslang.EOF And Not rslang.Bof Then
        SafeMoveFirst rslang
        On Error Resume Next
                
        With frm

            Do While Not rslang.EOF
                        
                If Not rslang.Fields.Item(sLanguage).value = vbNull Then
                    If Not rslang.Fields.Item(sLanguage).value = "" Then
                        If Not IsNull(rslang.Fields.Item("inx").value) Then 'IsArray(.Controls.Item(rslang.Fields.Item("Name").Value))
                            If Not rslang.Fields.Item(sLanguage).value = vbNull Then
                                If Not rslang.Fields.Item(sLanguage).value = "" Then
                                    Set ctr = .Controls.Item(rslang.Fields.Item("Name").value)(rslang.Fields.Item("inx").value)
                                    ctr.AutoSize = True
                                    ctr.caption = rslang.Fields.Item(sLanguage).value
                                    ctr.Text = rslang.Fields.Item(sLanguage).value
                                End If
                            End If

                        Else

                            If Not rslang.Fields.Item(sLanguage).value = vbNull Then
                                If Not rslang.Fields.Item(sLanguage).value = "" Then
                                    If Not rslang.Fields.Item("Type").value = "WinForm" Then
                                        .Controls.Item(rslang.Fields.Item("Name").value).caption = rslang.Fields.Item(sLanguage).value
                                        .Controls.Item(rslang.Fields.Item("Name").value).Text = rslang.Fields.Item(sLanguage).value
                                        .Controls.Item(rslang.Fields.Item("Name").value).AutoSize = True
                                    Else
                                        .caption = rslang.Fields.Item(sLanguage).value
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                    
                rslang.MoveNext
            
            Loop
                
            If bShow Then
                If Not lWHND = 0 Then
                    .Show iModal, lWHND
                Else
                    .Show
                End If
            End If
            
        End With
        
    End If
    
    rslang.Close
    
    Set rslang = Nothing
    
    Exit Sub
Hell:
    DebugPrint Err.Description
    Exit Sub
    
End Sub

Private Sub DumbStrings(Optional bOnlyForms As Boolean)
    
    Load frmAbout
    Load frmAboutOASIS
    Load frmAccessTest
    Load FrmAddBookMark
    Load frmAddIncident
    Load frmAddons
    Load frmAttributes
    
    Load frmChangeTracer
    Load frmChartSettings
    Load frmColorPalette
    
    Load frmDatePicker
    Load frmDownLoader
    Load frmDynamicContent
    
    Load frmEULA
    Load frmExportFormats
    
    Load frmFreqSettings
        
    Load frmLayerSelection
    Load frmLogin
    
    Load frmMain
    Load frmMapPrint
    Load frmMapproductPreview
    Load frmMapProductInfo
    Load frmMnuOASISProfile
    Load frmMnuOperations
    
    'Load frmOASISClientSynch
    Load frmOVMap
    
    Load frmScoring
    Load frmSecurityCharts
    Load frmSecChart
    
    Load frmTip
        
    Load frmWordBookMarks
    
    Load frmLanguagePicker

    Dim frmloaded As Form
    Dim i As Integer
    Dim sVal As String
    Dim rslang As New ADODB.Recordset

    rslang.Open "SELECT * FROM lang", m_Cnn, adOpenDynamic, adLockOptimistic

    With rslang

        For Each frmloaded In Forms
            
            On Error Resume Next
            
            If bOnlyForms Then
                .AddNew

                With .Fields
                    .Item("Type").value = "WinForm"
                    .Item("Container").value = frmloaded.Name
                    .Item("Name").value = frmloaded.Name
                    .Item("Default").value = frmloaded.caption
                End With

                .UpdateBatch
            Else
                DebugPrint "<" & frmloaded.Name & ">"
                
                'Loop through all form controls
        
                For i = 0 To frmloaded.Controls.Count - 1
            
                    '                With .Fields

                    'If TypeOf frmloaded.Controls(i) Is TextBox Then
                    '    .Item("Type").Value = ""
                    'ElseIf TypeOf frmloaded.Controls(i) Is CheckBox Then
                    '    .Item("Type").Value = ""
                    'ElseIf TypeOf frmloaded.Controls(i) Is Label Then
                    '    .Item("Type").Value = ""
                    'ElseIf TypeOf frmloaded.Controls(i) Is Frame Then
                    '    .Item("Type").Value = ""
                    'ElseIf TypeOf frmloaded.Controls(i) Is CommandButton Then
                    '    .Item("Type").Value = ""
                    'ElseIf TypeOf frmloaded.Controls(i) Is OptionButton Then
                    '    .Item("Type").Value = ""
                    'Else
                    
                    'End If
            
                    sVal = ""
                    sVal = frmloaded.Controls(i).caption

                    If Not Len(sVal) < 1 Then
                        sVal = frmloaded.Controls(i).Name & ".Caption =" & sVal
                    
                        .AddNew
                    
                        With .Fields
                            .Item("Type").value = TypeName(frmloaded.Controls(i))
                            .Item("Container").value = frmloaded.Name

                            'isArray
                            If frmloaded.Controls(i).Index >= 0 Then
                                .Item("inx").value = frmloaded.Controls(i).Index
                            End If

                            .Item("Name").value = frmloaded.Controls(i).Name
                            .Item("Default").value = frmloaded.Controls(i).caption
                        End With

                        .UpdateBatch

                    Else
                    
                        sVal = ""
                        sVal = frmloaded.Controls(i).Text
                        
                        If Not Len(sVal) < 1 Then
                            .AddNew
                        
                            With .Fields
                                sVal = frmloaded.Controls(i).Name & ".Text =" & sVal & frmloaded.Controls(i).Text
                        
                                .Item("Type").value = TypeName(frmloaded.Controls(i))
                                .Item("Container").value = frmloaded.Name

                                'isArray
                                If frmloaded.Controls(i).Index >= 0 Then
                                    .Item("inx").value = frmloaded.Controls(i).Index
                                End If

                                .Item("Name").value = frmloaded.Controls(i).Name
                                .Item("Default").value = frmloaded.Controls(i).Text
                        
                            End With
                                                
                            .UpdateBatch
                        End If
                    End If

                    DebugPrint sVal
                
                Next

            End If

        Next

    End With

    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next

End Sub

Private Sub ValidateInstall()
    Dim a_strArgs() As String
    Dim blnDebug As Boolean
    Dim strFileName As String
    Dim ofs As FileSystemObject
    Dim i As Integer
   
    a_strArgs = Split(Command$, " ")

    For i = LBound(a_strArgs) To UBound(a_strArgs)

        Select Case LCase(a_strArgs(i))

            Case "-q"
                DebugPrint a_strArgs(i + 1)
                Exit For

            Case "-path"
                Set ofs = New FileSystemObject
                On Error Resume Next

                If IsNumeric(a_strArgs(i + 1)) Then
                    g_sAppPath = (SpecialFolderPathEx(a_strArgs(i + 1)))
                
                    If Not ofs.FolderExists(g_sAppPath) Then
                        g_sAppPath = ""
                    End If
                
                Else

                    If ofs.FolderExists(a_strArgs(i + 1)) Then
                        g_sAppPath = a_strArgs(i + 1)
                    End If
                End If
                
                Set ofs = Nothing
        End Select
      
    Next i

    sListAllDrives

End Sub

Private Sub PathCommand()
    Dim a_strArgs() As String
    Dim blnDebug As Boolean
    Dim strFileName As String
    Dim mFileSysObj As FileSystemObject
    Dim i As Integer
   
    If InStr(Command$, "-a") > 0 Then
        a_strArgs = Split(Command$, """")

        For i = LBound(a_strArgs) To UBound(a_strArgs)

            Select Case LCase(Trim(a_strArgs(i)))

                Case "-a"

                    If InStr(a_strArgs(i + 1), ">>") Then
                        a_strArgs(i + 1) = Mid$(a_strArgs(i + 1), 1, InStr(a_strArgs(i + 1), ">>"))
                        a_strArgs(i + 1) = Replace(a_strArgs(i + 1), ">", "")
                    End If
            
                    g_sAppPath = Replace(a_strArgs(i + 1), """", "")
                    '>> separator should be used for server address if a different server is used. If left blank same server will be used
                    Exit For
            End Select
      
        Next i

    End If

End Sub

Private Sub CreateAppPath()

    Dim ofs As New FileSystemObject
    Dim i As Integer
    i = 1

    If Len(Command$) > 10 Then
        PathCommand
    Else
        While Environ$(i) <> ""

            If Mid(Environ$(i), 1, InStr(1, Environ(i), "=") - 1) = "oasisdocuments" Then
                g_sAppPath = Trim(right(Environ$(i), Len(Environ$(i)) - 1 - Len(Mid(Environ$(i), 1, InStr(1, Environ(i), "=") - 1))))
                g_sAppPath = g_sAppPath & "\iMMAP - OASIS\OASIS client"
            End If
    
            i = i + 1
        Wend
    End If

    If Not Len(g_sAppPath) > 1 Then

        If ofs.FolderExists(SpecialFolderPath(sftUserApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
            g_sAppPath = SpecialFolderPath(sftUserApplicationData) & "iMMAP - OASIS\OASIS client"
        ElseIf ofs.FolderExists(SpecialFolderPath(sftUserLocalApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
            g_sAppPath = SpecialFolderPath(sftUserLocalApplicationData) & "iMMAP - OASIS\OASIS client"
        ElseIf ofs.FolderExists(SpecialFolderPath(sftUserMyDocuments) & "iMMAP - OASIS\OASIS client\Data\DB") Then
            g_sAppPath = SpecialFolderPath(sftUserMyDocuments) & "iMMAP - OASIS\OASIS client"
        ElseIf ofs.FolderExists(SpecialFolderPath(sftCommonApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
            g_sAppPath = SpecialFolderPath(sftCommonApplicationData) & "iMMAP - OASIS\OASIS client"
        ElseIf ofs.FolderExists(SpecialFolderPath(sftCommonMyDocuments) & "iMMAP - OASIS\OASIS client\Data\DB") Then
            g_sAppPath = SpecialFolderPath(sftCommonMyDocuments) & "iMMAP - OASIS\OASIS client"
        End If

    End If
    
End Sub

Public Sub SaveStartUpParams()
        '<EhHeader>
        On Error GoTo SaveStartUpParams_Err
        '</EhHeader>
        Dim oIni As New clIniReader
    
100     If Not CreateNewINI(g_sAppPath & "\data\user\Sessions\sup.ini", oIni) Then Exit Sub
    
102     With oIni
104         .Path = g_sAppPath & "\data\user\Sessions\sup.ini"
106         .Section = "default"
            
            .Key = "DefServer"
            .value = frmLogin.ComServer.Text
            .AddKeyWithValue
            
            .Key = "RememberMe"
            .value = IIf(frmLogin.chkRememberUser.value = vbChecked, "true", "false")
            .AddKeyWithValue
            
            .Key = "Database"
            .value = IIf(bSQLServerInUse, "Microsoft SQL Server", "Microsoft Access Database")
            .AddKeyWithValue
            
108         .Key = "ApplicationSettings"
110         .value = g_udtSynchUpdateOptions.ApplicationSettings
112         .AddKeyWithValue
            
114         .Key = "AutoUpdate"
116         .value = g_udtSynchUpdateOptions.AutoUpdate
118         .AddKeyWithValue
            
120         .Key = "Charts"
122         .value = g_udtSynchUpdateOptions.Charts
124         .AddKeyWithValue

126         .Key = "DynamDataDefs"
128         .value = g_udtSynchUpdateOptions.DynamDataDefs
130         .AddKeyWithValue

132         .Key = "Feeds"
134         .value = g_udtSynchUpdateOptions.Feeds
136         .AddKeyWithValue

            '.Key = "ForceZero"
            '.Value = g_udtSynchUpdateOptions.ForceZero
            '.AddKeyWithValue
        
138         .Key = "GeoMarks"
140         .value = g_udtSynchUpdateOptions.GeoMarks
142         .AddKeyWithValue
        
144         .Key = "GISAttributeSettings"
146         .value = g_udtSynchUpdateOptions.GISAttributeSettings
148         .AddKeyWithValue
        
150         .Key = "Method"
152         .value = g_udtSynchUpdateOptions.lMethod
154         .AddKeyWithValue

156         .Key = "ManualSynchronisation"
158         .value = g_udtSynchUpdateOptions.ManualSynchronisation
160         .AddKeyWithValue

162         .Key = "MapProducts"
164         .value = g_udtSynchUpdateOptions.MapProducts
166         .AddKeyWithValue

168         .Key = "PrintTemplates"
170         .value = g_udtSynchUpdateOptions.PrintTemplates
172         .AddKeyWithValue

174         .Key = "SynchLayersSettings"
176         .value = g_udtSynchUpdateOptions.SynchLayersSettings
178         .AddKeyWithValue

180         .Key = "Thematics"
182         .value = g_udtSynchUpdateOptions.Thematics
184         .AddKeyWithValue

        End With

        '<EhFooter>
        Exit Sub

SaveStartUpParams_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.clientMain.SaveStartUpParams " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function CreateNewINI(StrIniFile As String, _
                             oINIReader As clIniReader) As Boolean
        '<EhHeader>
        On Error GoTo CreateNewINI_Err
        '</EhHeader>
        Dim fs As New FileSystemObject
    
100     If Not fs.FileExists(StrIniFile) Then

102         fs.CreateTextFile StrIniFile
104         Set fs = Nothing

106         With oINIReader
                .Path = StrIniFile
108             .Section = "default"
110             .Key = "starts"
112             .value = ""
114             .AddNewSection
            End With

        End If
    
116     CreateNewINI = True
        '<EhFooter>
        Exit Function

CreateNewINI_Err:
         
        '</EhFooter>
End Function

Private Sub CheckStartUpParams()
        '<EhHeader>
        On Error GoTo CheckStartUpParams_Err
        '</EhHeader>
        Dim oIni As New clIniReader
    
100     If Not CreateNewINI(g_sAppPath & "\data\user\Sessions\sup.ini", oIni) Then Exit Sub
    
102     With oIni
    
104         .Path = g_sAppPath & "\data\user\Sessions\sup.ini"
            
106         .Section = "default"
        
108         .Key = "ApplicationSettings"
110         g_udtSynchUpdateOptions.ApplicationSettings = CBool(IIf(.value = "", True, .value))
            
112         .Key = "AutoUpdate"
114         g_udtSynchUpdateOptions.AutoUpdate = CBool(IIf(.value = "", True, .value))
            
116         .Key = "Charts"
118         g_udtSynchUpdateOptions.Charts = CBool(IIf(.value = "", True, .value))

120         .Key = "DynamDataDefs"
122         g_udtSynchUpdateOptions.DynamDataDefs = CBool(IIf(.value = "", True, .value))

124         .Key = "Feeds"
126         g_udtSynchUpdateOptions.Feeds = CBool(IIf(.value = "", True, .value))

128         ' .Key = "ForceZero"
130         ' g_udtSynchUpdateOptions.ForceZero = IIf(.Value = "", False, CBool(.Value))
        
132         .Key = "GeoMarks"
134         g_udtSynchUpdateOptions.GeoMarks = CBool(IIf(.value = "", True, .value))
        
136         .Key = "GISAttributeSettings"
138         g_udtSynchUpdateOptions.GISAttributeSettings = CBool(IIf(.value = "", True, .value))
        
140         .Key = "Method"
142         g_udtSynchUpdateOptions.lMethod = CInt(IIf(.value = "", 0, .value))

144         .Key = "ManualSynchronisation"
146         g_udtSynchUpdateOptions.ManualSynchronisation = CBool(IIf(.value = "", False, .value))

148         .Key = "MapProducts"
150         g_udtSynchUpdateOptions.MapProducts = CBool(IIf(.value = "", True, .value))

152         .Key = "PrintTemplates"
154         g_udtSynchUpdateOptions.PrintTemplates = CBool(IIf(.value = "", True, .value))

156         .Key = "SynchLayersSettings"
158         g_udtSynchUpdateOptions.SynchLayersSettings = CBool(IIf(.value = "", True, .value))

160         .Key = "Thematics"
162         g_udtSynchUpdateOptions.Thematics = CBool(IIf(.value = "", True, .value))

            .Key = "Database"
            bSQLServerInUse = IIf(.value = "Microsoft SQL Server", True, False)

        End With
    
        '<EhFooter>
        Exit Sub

CheckStartUpParams_Err:
        MsgBox Err.desc
        Resume Next
        '</EhFooter>
End Sub

Sub Main()
        '<EhHeader>
        On Error GoTo Main_Err
        '</EhHeader>
        Dim lHwnd As Long
        
        Dim lngLocale As Long

        lngLocale = GetSystemDefaultLCID()

        If SetLocaleInfo(lngLocale, LOCALE_SSHORTDATE, "dd-MMM-yy") = False Then

            ' Handle error, possibly by writing it
            'MsgBox "date set error"
            ' to a server error log

        End If

100     Set m_oAES = New clsAES
102     Set m_frmDebug = New frmDebug
104     CreateAppPath
106     CheckIfEncryptedOK
        
108     g_sPermanentLyrs = ",oincidents,Draw Layer,Cosmetic,Buffers,"
        
110     If g_sAppPath = "" Then
112         g_sAppPath = App.Path
        End If
        
        'Load frmLogin
        'LoadLanguage "frmLogin", "Swedish", False, True
        'DumbStrings True
        'Exit Sub
                
114     If App.PrevInstance Then
            On Error Resume Next
116         MsgBox "Another Instance of OASIS Seems to be running. Please close the previous instance and try again.", vbInformation, "OASIS Client Start Up Routines"
            'get handle of previous instance
118         lHwnd = FindWindow(vbNullString, "m_frmMain")

            'show previous instance
120         If Not lHwnd = 0 Then
122             lHwnd = SetForegroundWindow(lHwnd)
            Else
124             lHwnd = FindWindow(vbNullString, "fLogin")

126             If Not lHwnd = 0 Then
128                 lHwnd = SetForegroundWindow(lHwnd)
                End If
            End If
            
130         End
        End If

132     With g_udtSynchUpdateOptions
134         .ApplicationSettings = True
136         .AutoUpdate = True
138         .Charts = True
140         .GeoMarks = True
142         .GISAttributeSettings = True
144         .ManualSynchronisation = False
146         .MapProducts = True
148         .PrintTemplates = True
150         .SynchLayersSettings = True
152         .Thematics = True
154         .Feeds = True
156         .DynamDataDefs = True
        End With

158     CheckStartUpParams
        
160     FixDB
        
        ' Exit Sub

        ' enable system themes
        Dim Themes As New VisualThemeAdapter
162     Themes.EnableVisualStyles

164     g_sAppSettingsTable = "AppSettings"
        
        'Checkif any command line arguments are passed.
        
166     ValidateInstall
        'Dim fLogin As frmLogin
        'Set fLogin = New frmLogin
                
168     If Len(Command$) > 10 Then
            If Not InStr(Command$, "-a") > 0 Then
170             Load frmLogin
            Else
                frmLogin.Show

                FormOnTop frmLogin
            End If

        Else
            'fLogin.Show vbModal
172         frmLogin.Show

174         FormOnTop frmLogin
            'Call FormOnTopEx(fLogin.hWnd, True)
            'fLogin.Show
            
            '106         If Not fLogin.OK Then
            '                'Login Failed so exit app
            '                On Error Resume Next
            '                Unload fLogin
            '108             End 'Err.Raise "666", "Login", "Login Failed"
            '            End If
            '
            '110         Unload fLogin
        End If
        
        '    frmClient.Show
        If frmLogin.txtUserName.Enabled And Len(frmLogin.txtUserName.Text) > 0 Then frmLogin.txtPassword.SetFocus
        '<EhFooter>
        Exit Sub

Main_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.clientMain.Main " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadMainWin()
        '<EhHeader>
        On Error GoTo LoadMainWin_Err
        Dim sConnString As String
        '</EhHeader>

100     'Load m_frmMain
        
102     'm_frmMain.WindowState = vbMaximized

        If m_Cnn.State = adStateClosed Then Exit Sub
        sConnString = m_Cnn.ConnectionString
        
        If Not bSQLServerInUse Then
            m_Cnn.Close
            m_Cnn.Open sConnString
        End If
        
        g_RSAppSettings.CursorLocation = g_sGlobalCursorLocation 'This was adUseServer
106     g_RSAppSettings.Open "SELECT * FROM " & g_sAppSettingsTable, m_Cnn, adOpenDynamic, adLockReadOnly
        
        g_RSLocalAppSettings.CursorLocation = g_sGlobalCursorLocation 'This was adUseServer
        g_RSLocalAppSettings.Open "SELECT * FROM InternalAppSettings", m_Cnn, adOpenDynamic, adLockOptimistic
        
105     'm_frmMain.Show
        
108     SafeMoveFirst g_RSAppSettings
110     g_RSAppSettings.Find "SettingName = 'InitMap'"
    
112     If bSQLServerInUse Then
            g_RSGISGridTableSettings.Open "SELECT * FROM GISgridTableSettings WHERE [visible] = 'True'", m_Cnn, adOpenDynamic, adLockBatchOptimistic
        Else
            g_RSGISGridTableSettings.Open "SELECT * FROM GISgridTableSettings WHERE [visible] = True", m_Cnn, adOpenDynamic, adLockBatchOptimistic
        End If
        
114     Set g_RSMaps = New ADODB.Recordset
    
116     g_RSMaps.Open "SELECT * FROM Maps", m_Cnn, adOpenDynamic, adLockReadOnly
    
118     m_frmMain.Init 'g_sAppPath & g_RSAppSettings.Fields.Item("SettingValue1").value '"\data\user\Maps\Iraq_Admin.TTKGP"
        m_frmMain.Show
        m_frmMain.GIS10.Lock

120     DoEvents

122     If Not g_sLanguage = "" And Not UCase$(g_sLanguage) = "DEFAULT" Then
124         m_frmMain.g_clsHotKey_HotKeyPress "INITWORKAROUND", MOD_ALT, vbKey0
        End If
        
        'Dim f As Form
        
        'DoEvents
        
        'If Not g_sLanguage = "" Then
        '    If Not m_Cnn.State = adStateClosed Then
            
        '    For Each f In Forms
        '        If f.Name = m_frmMain.Name Then
        '            LoadLanguage f.Name, g_sLanguage, m_Cnn
        '        End If
        '    Next

        '    End If
        'End If
        
        If Not g_DatabaseSpecExtent Is Nothing Then
            m_frmMain.GIS10.VisibleExtent = g_DatabaseSpecExtent
            m_frmMain.ctlZoomSlider1.SetZoomPointerFromExtent g_DatabaseSpecExtent
        End If

        m_frmMain.GIS10.Unlock

        '<EhFooter>
        Exit Sub

LoadMainWin_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.clientMain.LoadMainWin " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function StripNulls(OriginalStr As String) As String

    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If

    StripNulls = OriginalStr
End Function

Public Sub FixDB(Optional bForce As Boolean = False)
        '<EhHeader>
        On Error GoTo FixDB_Err
        '</EhHeader>
        
        If bSQLServerInUse Then Exit Sub
        
        Dim oDbUtility As New cMaintainDB
        Dim mFile As File, mFileSysObj As New FileSystemObject, mTxtStream As TextStream
        Dim iNumtimes As Integer
        Dim hSearch As Long ' Search Handle
        Dim WFD As WIN32_FIND_DATA
        Dim Cont As Integer
    
100     If bForce Then
102         If mFileSysObj.FileExists(g_sAppPath & "\data\db\OasisClient.mdb") Then
104             oDbUtility.CompactAccessDB g_sAppPath & "\data\db\OasisClient.mdb"
106             iNumtimes = 0
            Else
                Exit Sub
            End If

        Else
        
108         If Not mFileSysObj.FileExists(g_sAppPath & "\data\user\Sessions\main.dat") Then
110             mFileSysObj.CreateTextFile g_sAppPath & "\data\user\Sessions\main.dat", True
            
112             Open g_sAppPath & "\data\user\Sessions\main.dat" For Output As #1
114             Print #1, "0"
116             Close #1
    
            End If
        
118         Set mFile = mFileSysObj.GetFile(g_sAppPath & "\data\user\Sessions\main.dat")
120         Set mTxtStream = mFile.OpenAsTextStream(ForReading)
122         iNumtimes = CInt(mTxtStream.ReadLine)
        
124         Set mFile = Nothing
126         mTxtStream.Close
128         Set mTxtStream = Nothing
    
130         If iNumtimes > 10 Then
132             If MsgBox("OASIS has detected that you have not done any client Data maintainance for some time." & vbCrLf & "Would you like to do that now (this may take a few minutes)?", vbYesNo, "OASIS Database Maintainance") = vbYes Then
134                 oDbUtility.CompactAccessDB g_sAppPath & "\data\db\OasisClient.mdb"

136                 Cont = True
138                 hSearch = FindFirstFile(g_sAppPath & "\data\db\dynamicdata\*.mdb", WFD)

140                 If hSearch <> INVALID_HANDLE_VALUE Then
        
142                     Do While Cont
144                         oDbUtility.CompactAccessDB g_sAppPath & "\data\db\dynamicdata\" & StripNulls(WFD.cFileName)
146                         Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
                        Loop

                    End If

                End If
            
148             iNumtimes = 0
            Else
            
150             iNumtimes = iNumtimes + 1
            
            End If
        End If
    
152     Open g_sAppPath & "\data\user\Sessions\main.dat" For Output As #1
154     Print #1, CStr(iNumtimes)
156     Close #1
       
158     Set oDbUtility = Nothing
        '<EhFooter>
        Exit Sub

FixDB_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.clientMain.FixDB " & "at line " & Erl
        On Error Resume Next
        Set oDbUtility = Nothing
        '</EhFooter>
End Sub

Private Function FileExists(FullFileName As String) As Boolean

    On Error GoTo MakeF
    'If file does Not exist, there will be an Error
    Open FullFileName For Input As #1
    Close #1
    'no error, file exists
    FileExists = True
    Exit Function
MakeF:
    'error, file does Not exist
    FileExists = False
    Exit Function
End Function

Public Sub KILLEMHARD()
    On Error Resume Next
    DebugPrint ""
    Unload m_frmMain
    Set m_frmMain = Nothing
    
    Dim Process As Variant

    For Each Process In GetObject("winmgmts:").ExecQuery("select name from Win32_Process where name='OASISClient.exe'")
        Process.Terminate (0)
    Next
    
    End
End Sub

Public Sub CheckIfEncryptedOK()
        '<EhHeader>
        On Error GoTo CheckIfEncryptedOK_Err
        '</EhHeader>

        Dim fs As New FileSystemObject
        Dim oFile As Object
        Dim oTXTStream As TextStream
        Dim sKeyVals() As String
        Dim sKey As String
        Dim sPASS As String
        Dim bFileFound As Boolean
    
100     bFileFound = False
                            
102     If (fs.FileExists(g_sAppPath & "\data\user\Sessions\client.KEY")) = True Then
104         bFileFound = True
106         Set oTXTStream = fs.OpenTextFile(g_sAppPath & "\data\user\Sessions\client.KEY")
108         sKeyVals = Split(oTXTStream.ReadAll, vbCrLf)
110         sKey = sKeyVals(0)
112         sPASS = m_oAES.AESDecyptString(sKeyVals(1), sKey)
        End If
        
114     If bFileFound Then
            
116         g_sKey = sKey
118         g_bHasEncrypt = True
            DebugPrint "Server comms will be encrypted"
        
        Else
120         g_bHasEncrypt = False
            DebugPrint "Server comms will NOT be encrypted"
        End If
    
        '<EhFooter>
        Exit Sub

CheckIfEncryptedOK_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.clientMain.CheckIfEncryptedOK " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function CheckEncrypt(sValue As String) As String
        '<EhHeader>
        On Error GoTo CheckEncrypt_Err
        '</EhHeader>
            
100     If g_bHasEncrypt Then
102         CheckEncrypt = m_oAES.AESEncyptString(sValue, g_sKey)
        Else
104         CheckEncrypt = sValue
        End If

        '<EhFooter>
        Exit Function

CheckEncrypt_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.clientMain.CheckEncrypt " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function KeyGen(sKey As String) As String
        '<EhHeader>
        On Error GoTo KeyGen_Err
        '</EhHeader>
        Dim oMD5 As New clsMD5

100     KeyGen = oMD5.MD5(sKey)

        On Error Resume Next
102     Set oMD5 = Nothing
        '<EhFooter>
        Exit Function

KeyGen_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.clientMain.KeyGen " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

