Attribute VB_Name = "Module1"
Public sDynamDBPath As String
Private Type OPENFILENAME

        lStructSize As Long
        hWndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Public Const OFN_EXPLORER = &H80000 ' new look commdlg
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
'Folder
Private Type BrowseInfo
        hOwner As Long
        pidlRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfn As Long
        lParam As Long
        iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Const WM_USER = &H400
Public Const lPtr = (&H0 Or &H40)
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

'Open/Save
Private Declare Function GetSaveFileName Lib "COMDLG32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long



Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Public Const LF_FACESIZE = 32

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Declare Function SHGetSpecialFolderLocation _
                Lib "shell32.dll" (ByVal hWndOwner As Long, _
                                   ByVal nFolder As Long, _
                                   pidl As ITEMIDLIST) As Long



'''''''''''
Private Type SHELLITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHELLITEMID
End Type

Public Enum SpecialFolderTypes
    sftCDBurningCache = 59&
    sftCommonAdminTools = 47&
    sftCommonApplicationData = 35&
    sftCommonDesktop = 25&
    sftCommonDocumentTemplates = 45&
    sftCommonFavorites = 31&
    sftCommonMyDocuments = 46&
    sftCommonMyPictures = 54&
    sftCommonProgramFiles = 43&
    sftCommonStartMenu = 22&
    sftCommonStartMenuPrograms = 23&
    sftCommonStartup = 24&
    sftFonts = 20&
    sftProgramFiles = 38&
    sftSystem32Folder = 41&
    sftSystemFolder = 37&
    sftThemes = 56&
    sftUserAdminTools = 48&
    sftUserApplicationData = 26&
    sftUserCookies = 33&
    sftUserDesktop = 16&
    sftUserDocumentTemplates = 21&
    sftUserFavorites = 6&
    sftUserHistory = 34&
    sftUserLocalApplicationData = 28&
    sftUserMyDocuments = 5&
    sftUserMyMusic = 13&
    sftUserMyPictures = 39&
    sftUserNetHood = 19&
    sftUserPrintHood = 27&
    sftUserProfileFolder = 40&
    sftUserRecentDocuments = 8&
    sftUserSendTo = 9&
    sftUserStartMenu = 11&
    sftUserStartMenuPrograms = 2&
    sftUserStartup = 7&
    sftUserTempInternetFiles = 32&
    sftWindowsFolder = 36&
End Enum


Private m_cHookedDialog As Long

Property Let HookedDialog(ByRef cThis As GCommonDialog)
    'Set cHookedDialog = cThis
    m_cHookedDialog = ObjPtr(cThis)
End Property
Property Get HookedDialog() As GCommonDialog
    Dim oT As GCommonDialog

    If (m_cHookedDialog <> 0) Then
        ' Turn the pointer into an illegal, uncounted interface
        CopyMemory oT, m_cHookedDialog, 4
        ' Do NOT hit the End button here! You will crash!
        ' Assign to legal reference
        Set HookedDialog = oT
        ' Still do NOT hit the End button here! You will still crash!
        ' Destroy the illegal reference
        CopyMemory oT, 0&, 4
    End If

End Property
Public Function SpecialFolderPath(ByVal lngFolderType As SpecialFolderTypes) As String

    Dim strPath As String
    Dim IDL As ITEMIDLIST
    Dim MAX_PATH As Integer
    MAX_PATH = 255

    SpecialFolderPath = ""

    If SHGetSpecialFolderLocation(0&, lngFolderType, IDL) = 0& Then
        strPath = Space$(MAX_PATH)

        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath) Then
            SpecialFolderPath = Left$(strPath, InStr(strPath, vbNullChar) - 1&) & "\"
        End If
    End If

End Function
Public Sub ClearHookedDialog()
    m_cHookedDialog = 0
End Sub

Public Function DialogHookFunction(ByVal hDlg As Long, _
                                   ByVal msg As Long, _
                                   ByVal wParam As Long, _
                                   ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        DialogHookFunction = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function

Public Function PrintHookProc(ByVal hDlg As Long, _
                              ByVal msg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        PrintHookProc = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function

Public Function PrintSetupHookProc(ByVal hDlg As Long, _
                                   ByVal msg As Long, _
                                   ByVal wParam As Long, _
                                   ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        PrintSetupHookProc = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function

Public Function PageSetupHook(ByVal hDlg As Long, _
                              ByVal msg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        PageSetupHook = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function

Public Function CCHookProc(ByVal hDlg As Long, _
                           ByVal msg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        CCHookProc = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function

Public Function CFHookProc(ByVal hDlg As Long, _
                           ByVal msg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        CFHookProc = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function


Public Function SpecialFolderPathEx(ByVal lngFolderType As Long) As String

Dim strPath As String
Dim IDL As ITEMIDLIST
Dim MAX_PATH As Integer
MAX_PATH = 255

SpecialFolderPathEx = ""

If SHGetSpecialFolderLocation(0&, lngFolderType, IDL) = 0& Then
    strPath = Space$(MAX_PATH)

    If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath) Then
        SpecialFolderPathEx = Left$(strPath, InStr(strPath, vbNullChar) - 1&) & "\"
    End If
End If

End Function


Public Sub CreateDynamDBPath()
        '<EhHeader>
        On Error GoTo CreateDynamDBPath_Err
        '</EhHeader>

        Dim ofs As New FileSystemObject

100     If ofs.FolderExists(SpecialFolderPath(sftUserApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
102         sDynamDBPath = SpecialFolderPath(sftUserApplicationData) & "iMMAP - OASIS\OASIS client\Data\dynamicdata"
104     ElseIf ofs.FolderExists(SpecialFolderPath(sftUserLocalApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
106         sDynamDBPath = SpecialFolderPath(sftUserLocalApplicationData) & "iMMAP - OASIS\OASIS client\Data\dynamicdata"
108     ElseIf ofs.FolderExists(SpecialFolderPath(sftUserMyDocuments) & "iMMAP - OASIS\OASIS client\Data\DB") Then
110         sDynamDBPath = SpecialFolderPath(sftUserMyDocuments) & "iMMAP - OASIS\OASIS client\Data\dynamicdata"
112     ElseIf ofs.FolderExists(SpecialFolderPath(sftCommonApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
114         sDynamDBPath = SpecialFolderPath(sftCommonApplicationData) & "iMMAP - OASIS\OASIS client\Data\dynamicdata"
116     ElseIf ofs.FolderExists(SpecialFolderPath(sftCommonMyDocuments) & "iMMAP - OASIS\OASIS client\Data\DB") Then
118         sDynamDBPath = SpecialFolderPath(sftCommonMyDocuments) & "iMMAP - OASIS\OASIS client\Data\dynamicdata"
        End If
    
        '<EhFooter>
        Exit Sub

CreateDynamDBPath_Err:
        MsgBox Err.Description & vbCrLf & "in Project2.Module1.CreateDynamDBPath " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function FileExists(sFullPath As String) As Boolean
        '<EhHeader>
        On Error GoTo FileExists_Err
        '</EhHeader>
        Dim oFile As New Scripting.FileSystemObject
100     FileExists = oFile.FileExists(sFullPath)
        '<EhFooter>
        Exit Function

FileExists_Err:
        MsgBox Err.Description & vbCrLf & _
               "in Project2.Module1.FileExists " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function VerifyDatabasePath(sPassedDatabasePath As String) As Boolean
        '<EhHeader>
        On Error GoTo VerifyDatabasePath_Err
        '</EhHeader>

100     VerifyDatabasePath = False
102     sDatabaseName = sPassedDatabasePath
    
104     Do Until InStr(1, sDatabaseName, "\", vbTextCompare) = 0
106         sDatabaseName = Right(sDatabaseName, Len(sDatabaseName) - InStr(1, sDatabaseName, "\", vbTextCompare))
        Loop
    
108     sDatabasePath = Left$(sPassedDatabasePath, Len(sPassedDatabasePath) - Len(sDatabaseName) - 1)
    
110     If UCase(sDatabasePath) = UCase(sDynamDBPath) Then
    
            'If MsgBox("Do you want to save this dynamic database definition?", vbYesNo, "Save?") = vbYes Then
    
            'RSDefinedDynDatabases.AddNew Array("DynamDBDefGUID", "GroupGUID", "DatabaseName", "DatabasePath", "ExcludedFields"), Array(GUIDGen, RSLocalUserGroups!sGUID, sDatabaseName, sDatabasePath, sExcludedFields)
112         VerifyDatabasePath = True
        
            'End If
        
114         If Not FileExists(sPassedDatabasePath) Then
116             VerifyDatabasePath = False
118             MsgBox "This database does not exist: " & sPassedDatabasePath, vbCritical, "Operation Aborted"
            End If
    
        Else
    
120         MsgBox "Operation cancelled.  You must use the path: " & sDynamDBPath, vbCritical, "Error!"
    
        End If
    
        'MsgBox PromptDatabaseDefinitionSave

        '<EhFooter>
        Exit Function

VerifyDatabasePath_Err:
        MsgBox Err.Description & vbCrLf & "in Project2.Module1.VerifyDatabasePath " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function


Private Function BrowseCallbackProcStr(ByVal hwnd As Long, _
                                       ByVal uMsg As Long, _
                                       ByVal lParam As Long, _
                                       ByVal lpData As Long) As Long

    If uMsg = 1 Then
        Call SendMessage(hwnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End If

End Function

Public Function BrowseForFolder(strTitle As String, _
                                lngHwnd As Long, _
                                Optional strInitialDirectory As String) As String

    Dim Browse_for_folder As BrowseInfo
    Dim lngItemID As Long
    Dim lngInitDirPointer As Long
    Dim strTempPath As String * 256

    If strInitialDirectory = "" Then strInitialDirectory = g_sAppPath

    With Browse_for_folder
        .hOwner = lngHwnd 'Window Handle
        .lpszTitle = strTitle 'Dialog Title
        .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr) 'Dialog callback function that preselectes the folder specified
        lngInitDirPointer = LocalAlloc(lPtr, Len(strInitialDirectory) + 1) 'Allocate a string
        Call CopyMemory(ByVal lngInitDirPointer, ByVal strInitialDirectory, Len(strInitialDirectory) + 1) 'Copy the path to the string
        .lParam = lngInitDirPointer  'The folder to preselect
    End With

    lngItemID = SHBrowseForFolder(Browse_for_folder) 'Execute the BrowseForFolder API

    If lngItemID Then
        If SHGetPathFromIDList(lngItemID, strTempPath) Then ' Get the path for the selected folder in the dialog
            BrowseForFolder = Left$(strTempPath, InStr(strTempPath, vbNullChar) - 1) ' Take only the path without the nulls
        End If

        Call CoTaskMemFree(lngItemID) 'Free the lngItemID
    End If

    Call LocalFree(lngInitDirPointer) 'Free the string from the memory

End Function

Private Function FunctionPointer(FunctionAddress As Long) As Long

    FunctionPointer = FunctionAddress

End Function

'"JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|CompuServe GIF (*.gif)|*.gif"
Public Function OpenDialog(strFilter As String, _
                           strTitle As String, _
                           strDefaultExtension As String, _
                           strInitialDirectory As String, _
                           lngHwnd As Long) As String

    On Error GoTo Problems
    Dim OpenFile As OPENFILENAME
    Dim strTemp As String
    Dim intNull As Integer

    If Right$(strFilter, 1) <> Chr$(0) Then strFilter = strFilter & Chr$(0)
    strFilter = Replace(strFilter, "|", Chr$(0))

    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hWndOwner = lngHwnd
    OpenFile.lpstrInitialDir = strInitialDirectory
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = strFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String$(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrTitle = strTitle
    OpenFile.lpstrDefExt = strDefaultExtension
    OpenFile.Flags = OFN_HIDEREADONLY

    If GetOpenFileName(OpenFile) = 0 Then
        OpenDialog = ""
    Else
        strTemp = OpenFile.lpstrFile
        intNull = InStr(1, strTemp, Chr$(0))
        OpenDialog = Mid$(strTemp, 1, intNull - 1)
    End If

    Exit Function

Problems:
    MsgBox Err.Description, 16, "Error " & Err.Number

End Function

Public Function SaveDialog(strFilter As String, _
                           strTitle As String, _
                           strInitialDirectory As String, _
                           lngHwnd As Long, _
                           Optional strFileName As String) As String

    Dim OpenFile As OPENFILENAME
    Dim strExtension As String

    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hWndOwner = lngHwnd
    OpenFile.hInstance = App.hInstance

    If Right$(strFilter, 1) <> "|" Then strFilter = strFilter + "|"

    strFilter = Replace(strFilter, "|", Chr$(0))

    If strFileName = "" Then strFileName = Space$(254) Else strFileName = strFileName & Space$(254 - Len(strFileName))

    OpenFile.lpstrFilter = strFilter
    OpenFile.lpstrFile = strFileName
    OpenFile.nMaxFile = 255
    OpenFile.lpstrFileTitle = Space$(254)
    OpenFile.nMaxFileTitle = 255
    OpenFile.lpstrInitialDir = strInitialDirectory
    OpenFile.lpstrTitle = strTitle
    OpenFile.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT

    If GetSaveFileName(OpenFile) Then
        SaveDialog = Trim$(OpenFile.lpstrFile)
        strExtension = Mid$(Right$(strFilter, 5), 1, 4)
        strFileName = Left$(SaveDialog, Len(SaveDialog) - 1)

        If Right$(strFileName, 4) = strExtension Then strExtension = ""
        SaveDialog = strFileName & strExtension

        If strFilter = "*.*" & Chr$(0) Then SaveDialog = strFileName
    Else
        SaveDialog = ""
    End If

End Function




