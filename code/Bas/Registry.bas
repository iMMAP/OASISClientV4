Attribute VB_Name = "Registry"
Option Explicit


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_INVALID_PARAMETER = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const ERROR_MORE_DATA = 234

Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_OPTION_NON_VOLATILE = 0

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Enum RegistryLTypes
    REG_SZ = 1
    REG_BINARY = 3
    REG_DWORD = 4
End Enum

Declare Function RegDeleteValue _
        Lib "advapi32" _
        Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                 ByVal lpValueName As String) As Long
Declare Function RegEnumValue _
        Lib "advapi32" _
        Alias "RegEnumValueA" (ByVal hKey As Long, _
                               ByVal dwIndex As Long, _
                               ByVal lpValueName As String, _
                               lpcbValueName As Long, _
                               ByVal lpReserved As Long, _
                               lpType As Long, _
                               ByVal lpData As String, _
                               lpcbData As Long) As Long
Declare Function RegEnumKeyEx _
        Lib "advapi32" _
        Alias "RegEnumKeyExA" (ByVal hKey As Long, _
                               ByVal dwIndex As Long, _
                               ByVal lpName As String, _
                               lpcbName As Long, _
                               lpReserved As Long, _
                               ByVal lpClass As String, _
                               lpcbClass As Long, _
                               lpftLastWriteTime As FILETIME) As Long

Declare Function RegOpenKeyEx _
        Lib "advapi32" _
        Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                               ByVal lpSubKey As String, _
                               ByVal ulOptions As Long, _
                               ByVal samDesired As Long, _
                               phkResult As Long) As Long
Declare Function RegCloseKey _
        Lib "advapi32" (ByVal hKey As Long) As Long

Declare Function RegCreateKeyEx _
        Lib "advapi32" _
        Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                 ByVal lpSubKey As String, _
                                 ByVal Reserved As Long, _
                                 ByVal lpClass As String, _
                                 ByVal dwOptions As Long, _
                                 ByVal samDesired As Long, _
                                 ByVal lpSecurityAttributes As Long, _
                                 phkResult As Long, _
                                 lpdwDisposition As Long) As Long
Declare Function RegDeleteKey _
        Lib "advapi32" _
        Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                               ByVal lpSubKey As String) As Long

Declare Function RegQueryValueExString _
        Lib "advapi32" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                  ByVal lpValueName As String, _
                                  ByVal lpReserved As Long, _
                                  lpType As Long, _
                                  ByVal lpData As String, _
                                  lpcbData As Long) As Long
Declare Function RegQueryValueExLong _
        Lib "advapi32" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                  ByVal lpValueName As String, _
                                  ByVal lpReserved As Long, _
                                  lpType As Long, _
                                  lpData As Long, _
                                  lpcbData As Long) As Long
Declare Function RegQueryValueExNULL _
        Lib "advapi32" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                  ByVal lpValueName As String, _
                                  ByVal lpReserved As Long, _
                                  lpType As Long, _
                                  ByVal lpData As Long, _
                                  lpcbData As Long) As Long
Declare Function RegQueryValueEx _
        Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                  ByVal lpValueName As String, _
                                  ByVal lpReserved As Long, _
                                  lpType As Long, _
                                  lpData As Any, _
                                  lpcbData As Long) As Long

Declare Function RegSetValueExString _
        Lib "advapi32" _
        Alias "RegSetValueExA" (ByVal hKey As Long, _
                                ByVal lpValueName As String, _
                                ByVal Reserved As Long, _
                                ByVal dwType As Long, _
                                ByVal lpValue As String, _
                                ByVal cbData As Long) As Long
Declare Function RegSetValueExLong _
        Lib "advapi32" _
        Alias "RegSetValueExA" (ByVal hKey As Long, _
                                ByVal lpValueName As String, _
                                ByVal Reserved As Long, _
                                ByVal dwType As Long, _
                                lpValue As Long, _
                                ByVal cbData As Long) As Long
Public Declare Function RegSetValueEx _
               Lib "advapi32.dll" _
               Alias "RegSetValueExA" (ByVal hKey As Long, _
                                       ByVal lpValueName As String, _
                                       ByVal Reserved As Long, _
                                       ByVal dwType As Long, _
                                       lpData As Any, _
                                       ByVal cbData As Long) As Long
Declare Function ShellExecute _
        Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As Long, _
                               ByVal lpOperation As String, _
                               ByVal lpFile As String, _
                               ByVal lpParameters As String, _
                               ByVal lpDirectory As String, _
                               ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Function CreateNewKey(MainKey As Long, SubKey As String)

    Dim hNewKey As Long
    Dim lRetVal As Long

    On Error GoTo Problems
    RegCreateKeyEx MainKey, SubKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal
    RegCloseKey (hNewKey)

    Exit Function

Problems:
    MsgBox Err.Description & " (CreateNewKey)", vbExclamation, "Error number " & Err.Number

End Function

Function DeleteKey(MainKey As Long, SubKey As String)

    On Error GoTo Problems
    RegDeleteKey MainKey, SubKey

    Exit Function

Problems:
    MsgBox Err.Description & " (DeleteKey)", vbExclamation, "Error number " & Err.Number

End Function

Function DeleteValue(MainKey As Long, SubKey As String, ValueName As String)

    Dim hKey As Long

    On Error GoTo Problems
    RegOpenKeyEx MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey
    RegDeleteValue hKey, ValueName
    RegCloseKey (hKey)

    Exit Function

Problems:
    MsgBox Err.Description & " (DeleteValue)", vbExclamation, "Error number " & Err.Number

End Function

Function KeyCount(MainKey As Long, SubKey As String)

    Dim ft As FILETIME
    Dim hKey As Long
    Dim Res As Long
    Dim Count As Long
    Dim keyname As String, classname As String
    Dim KeyLen As Long, classlen As Long

    On Error GoTo Problems
    RegOpenKeyEx MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey

    Do
        KeyLen = 2000
        classlen = 2000
        keyname = Space$(KeyLen)
        classname = Space$(classlen)
        Res = RegEnumKeyEx(hKey, Count, keyname, KeyLen, 0, classname, classlen, ft)
        Count = Count + 1
    Loop While Res = 0

    KeyCount = Count - 1
    RegCloseKey (hKey)

    Exit Function

Problems:
    MsgBox Err.Description & " (KeyCount)", vbExclamation, "Error number " & Err.Number

End Function

Function KeyExists(MainKey As Long, SubKey As String) As Boolean

    Dim hKey As Long

    On Error GoTo Problems

    If RegOpenKeyEx(MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey) = 0 Then RegCloseKey hKey: KeyExists = True Else KeyExists = False

    Exit Function

Problems:
    MsgBox Err.Description & " (KeyExists)", vbExclamation, "Error number " & Err.Number

End Function

Function QueryValue(MainKey As Long, SubKey As String, ValueName As String, lType As RegistryLTypes)

    Dim hKey As Long
    Dim vValue

    RegOpenKeyEx MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey
    QueryValueEx hKey, ValueName, vValue, lType
    RegCloseKey (hKey)
    QueryValue = vValue

End Function

Private Function QueryValueEx(ByVal lhKey As Long, _
                              ByVal szValueName As String, _
                              vValue As Variant, _
                              lType As RegistryLTypes) As Variant

    Dim cch As Long
    Dim lrc As Long
    Dim lValue As Long
    Dim sValue As String

    ReDim bData(0) As Byte

    On Error GoTo Problems

    Select Case lType

        Case REG_SZ

            lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
            sValue = String$(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)

            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If

        Case REG_DWORD

            lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)

            If lrc = ERROR_NONE Then vValue = lValue

        Case REG_BINARY

            lrc = RegQueryValueEx(lhKey, szValueName, 0&, lType, bData(0), cch)

            If lrc = ERROR_NONE Or lrc = ERROR_MORE_DATA Then

                ReDim bData(0 To cch - 1)
                lrc = RegQueryValueEx(lhKey, szValueName, CLng(0), lType, bData(0), cch)
            End If

            vValue = bData

        Case Else
            lrc = -1
    End Select

QueryValueExExit:

    If Right$(vValue, 1) = Chr$(0) Then
        vValue = Left$(vValue, Len(vValue) - 1)
    End If

    QueryValueEx = vValue

    Exit Function

Problems:
    MsgBox Err.Description & " (QueryValue)", vbExclamation, "Error number " & Err.Number
    Resume QueryValueExExit

End Function

Function SetKeyValue(MainKey As Long, SubKey As String, ValueName As String, ValueSetting As Variant, lType As RegistryLTypes)

    Dim lValue As Long
    Dim sValue As String
    Dim hKey As Long
    Dim lLength As Long
    Dim i As Integer

    ReDim bData(0) As Byte

    On Error GoTo Problems
    RegOpenKeyEx MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey

    Select Case lType

        Case REG_SZ
            sValue = ValueSetting & Chr$(0)
            RegSetValueExString hKey, ValueName, 0&, lType, sValue, Len(sValue)

        Case REG_DWORD
            lValue = ValueSetting
            RegSetValueExLong hKey, ValueName, 0&, lType, lValue, 4

        Case REG_BINARY ' Free form binary

            lLength = (UBound(ValueSetting) - LBound(ValueSetting)) + 1
            ReDim bData(LBound(ValueSetting) To UBound(ValueSetting))

            For i = LBound(ValueSetting) To UBound(ValueSetting)
                bData(i) = CByte(ValueSetting(i))
            Next i

            RegSetValueEx hKey, ValueName, 0&, lType, bData(LBound(ValueSetting)), lLength

    End Select

    RegCloseKey (hKey)

    Exit Function

Problems:
    MsgBox Err.Description & " (SetKeyValue)", vbExclamation, "Error number " & Err.Number

End Function

Function ValueCount(MainKey As Long, SubKey As String)

    Dim hKey As Long
    Dim Res As Long
    Dim Count As Long
    Dim lType As Long
    Dim ValueName As String, Valuelen As Long
    Dim lData As String, Datalen As Long

    On Error GoTo Problems
    RegOpenKeyEx MainKey, SubKey, 0, KEY_ALL_ACCESS, hKey

    Do
        ValueName = Space$(255)
        Valuelen = Len(ValueName)
        lData = Space$(255)
        Datalen = Len(lData)
        Res = RegEnumValue(hKey, Count, ValueName, Valuelen, 0, lType, lData, Datalen)
        Count = Count + 1
    Loop While Res = 0

    ValueCount = Count
    RegCloseKey (hKey)

    Exit Function

Problems:
    MsgBox Err.Description & " (ValueCount)", vbExclamation, "Error number " & Err.Number

End Function

