Attribute VB_Name = "modFileAssoc"
' IMPORTANT NOTICE: Applying a new file association will only be
' effective AFTER computer restart !!!
'
'
' >>> Make an association
'
' MakeFileAssociation(Extension, PathToApplication, ApplicationName, Description, Optional FullIconPath)
'
' Where:
'
' Extension (string)         Your new filetypes extension (without the point !)
' PathToApplication (string) The full path (without exe name) of your program
' ApplicationName (string)   The exe name (including .exe)
' Description (string)       Description of the filetype as shown in the explorer
' FullIconPath (string)      Full path of the associated icon (including .ico)
'
'
' >>> Delete an association
'
' DeleteFileAssociation(Extension)
'
' Where:
'
' Extension (string) The filetypes extension you want to delete(without the point !)
'
'
' >>> Check if an association exists
'
' return = CheckFileAssociation(Extension)
'
' Where:
'
' Extension (string) The filetypes extension you want to verify
' return    (string) the name of the associated exe, or empty
'                    if no associated exe-file exists
'

Private Declare Function RegCreateKey _
                Lib "advapi32.dll" _
                Alias "RegCreateKeyA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       phkResult As Long) As Long
Private Declare Function RegSetValue _
                Lib "advapi32.dll" _
                Alias "RegSetValueA" (ByVal hKey As Long, _
                                      ByVal lpSubKey As String, _
                                      ByVal dwType As Long, _
                                      ByVal lpData As String, _
                                      ByVal cbData As Long) As Long
Private Declare Function RegCloseKey _
                Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       phkResult As Long) As Long
Private Declare Function RegQueryValue _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          ByVal lpData As String, _
                                          lpcbData As Long) As Long
Private Declare Function RegDeleteKey _
                Lib "advapi32.dll" _
                Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String) As Long

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&

Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_SET_VALUE = &H2&
Private Const MAX_PATH = 260&
Private Const REG_DWORD As Long = 4
Private Const REG_SZ = 1
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL

Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY

Private Declare Function fCreateShellLink Lib "VB6STKIT.DLL" (ByVal _
        lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
        lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

'To update windows Icon Cache
Private Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, _
                        ByVal uFlags As Long, ByVal dwItem1 As Long, _
                        ByVal dwItem2 As Long)
                        
' A file type association has changed.
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0

'Public Const HKEY_CLASSES_ROOT = &H80000000
'Public Const HKEY_CURRENT_USER = &H80000001
'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const HKEY_USERS = &H80000003

Public Sub MakeFileAssociation(Extension As String, _
                               PathToApplication As String, _
                               ApplicationName As String, _
                               Description As String, _
                               Optional FullIconPath As String)
        '<EhHeader>
        On Error GoTo MakeFileAssociation_Err
        '</EhHeader>
        Dim Ret&

100     If Left(PathToApplication, 1) <> "\" Then PathToApplication = PathToApplication & "\"
        'Create a Root entry called .XXX associated with application name
102     sKeyName = "." & Extension
104     sKeyValue = ApplicationName
106     Ret& = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)
        'Set application key and file description
108     sKeyName = ApplicationName
110     sKeyValue = Description
112     Ret& = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)

        'This sets the default icon for XXX_auto_file
114     If FullIconPath <> "" Then
116         sKeyName = ApplicationName & "\DefaultIcon"
118         sKeyValue = FullIconPath & ",0"
120         Ret& = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)
        End If

        'This sets the command line for XXX_auto_file
122     sKeyName = ApplicationName & "\shell\open\command"
124     sKeyValue = Chr(34) & PathToApplication & ApplicationName & ".exe" & Chr(34) & " %1"
126     Ret& = WriteKey(HKEY_CLASSES_ROOT, sKeyName, "", sKeyValue)
        '<EhFooter>
        Exit Sub

MakeFileAssociation_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.modFileAssoc.MakeFileAssociation " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub DeleteFileAssociation(Extension As String)
        '<EhHeader>
        On Error GoTo DeleteFileAssociation_Err
        '</EhHeader>
        Dim Application As String
        Dim Ret&
        'check if filetype is registred
100     Application = ReadKey(HKEY_CLASSES_ROOT, "." & Extension, "", "")

102     If Application <> "" Then
            'delete file extension
104         Ret& = DeleteKey(HKEY_CLASSES_ROOT, "." & Extension)
            'delete command lines
106         Ret& = DeleteKey(HKEY_CLASSES_ROOT, Application)
        End If

        '<EhFooter>
        Exit Sub

DeleteFileAssociation_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.modFileAssoc.DeleteFileAssociation " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function CheckFileAssociation(ByVal Extension As String) As String
        '<EhHeader>
        On Error GoTo CheckFileAssociation_Err
        '</EhHeader>
100     Extension = "." & Extension
        'read in the program name associated with this filetype
102     CheckFileAssociation = ReadKey(HKEY_CLASSES_ROOT, Extension, "", "")
        '<EhFooter>
        Exit Function

CheckFileAssociation_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.modFileAssoc.CheckFileAssociation " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function ReadKey(ByVal KeyName As String, _
                        ByVal SubKeyName As String, _
                        ByVal ValueName As String, _
                        ByVal DefaultValue As String) As String
        '<EhHeader>
        On Error GoTo ReadKey_Err
        '</EhHeader>
        Dim sBuffer As String
        Dim lBufferSize As Long
        Dim Ret&
100     sBuffer = Space(255)
102     lBufferSize = Len(sBuffer)
104     Ret& = RegOpenKey(KeyName, SubKeyName, 0, KEY_READ, lphKey&)

106     If Ret& = ERROR_SUCCESS Then
108         Ret& = RegQueryValue(lphKey&, ValueName, 0, REG_SZ, sBuffer, lBufferSize)
110         Ret& = RegCloseKey(lphKey&)
        Else
112         Ret& = RegCloseKey(lphKey&)
        End If

114     sBuffer = Trim(sBuffer)

116     If sBuffer <> "" Then
118         sBuffer = Left(sBuffer, Len(sBuffer) - 1)
        Else
120         sBuffer = DefaultValue
        End If

122     ReadKey = sBuffer
        '<EhFooter>
        Exit Function

ReadKey_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.modFileAssoc.ReadKey " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function WriteKey(ByVal KeyName As String, _
                         ByVal SubKeyName As String, _
                         ByVal ValueName As String, _
                         ByVal KeyValue As String) As Long
        '<EhHeader>
        On Error GoTo WriteKey_Err
        '</EhHeader>
        Dim Ret&
100     Ret& = RegCreateKey&(KeyName, SubKeyName, lphKey&)

102     If Ret& = ERROR_SUCCESS Then
104         Ret& = RegSetValue&(lphKey&, ValueName, REG_SZ, KeyValue, 0&)
        Else
106         Ret& = RegCloseKey(lphKey&)
        End If

108     WriteKey = Ret&
        '<EhFooter>
        Exit Function

WriteKey_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.modFileAssoc.WriteKey " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function DeleteKey(ByVal KeyName As String, _
                          ByVal SubKeyName As String) As Long
        '<EhHeader>
        On Error GoTo DeleteKey_Err
        '</EhHeader>
        Dim Ret&
100     Ret& = RegOpenKey(KeyName, SubKeyName, 0, KEY_WRITE, lphKey&)

102     If Ret& = ERROR_SUCCESS Then
104         Ret& = RegDeleteKey(lphKey&, "") 'delete the key
106         Ret& = RegCloseKey(lphKey&)
        End If

108     DeleteKey = Ret&
        '<EhFooter>
        Exit Function

DeleteKey_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.modFileAssoc.DeleteKey " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

