Attribute VB_Name = "fileDirFunctions"
Option Explicit

Private Const HKEY_CLASSES_ROOT = &H80000000

Private Declare Function RegCreateKey Lib _
   "advapi32.dll" Alias "RegCreateKeyA" _
   (ByVal hKey As Long, ByVal lpSubKey As _
   String, phkResult As Long) As Long

Private Declare Function RegCloseKey Lib _
   "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegSetValueEx Lib _
   "advapi32.dll" Alias "RegSetValueExA" _
   (ByVal hKey As Long, ByVal _
   lpValueName As String, ByVal _
   Reserved As Long, ByVal dwType _
   As Long, lpData As Any, ByVal _
   cbData As Long) As Long

Private Const REG_SZ = 1

Function StripFileName(FilePath As String) As String

    Dim Path As Variant
    Path = Split(FilePath, "\")
    StripFileName = Path(UBound(Path))
End Function

Public Function CopyFile(srcFile As String, dstFile As String)

    'this copies a file byte-for-byte
    'or you could just use good old FileCopy
    '     :-)
    On Error Resume Next 'If we Get an error, keep going
    Dim Copy As Long
    Dim CopyByteForByte As Byte 'the variables
    
    Open srcFile For Binary Access Write As #1 'open the destination file so we can write To it
    Open dstFile For Binary Access Read As #2 'open the source file so we can read from it


    For Copy = 1 To LOF(2) 'Copy The SourceFile Byte-For-Byte
        Put #1, , CopyByteForByte 'Put the byte In the destination file
    Next Copy 'stop the Loop

    
End Function

Public Function CreateFileAssociation(AppName As String, _
ByVal AppExtension As String, AppCommand As String) As Boolean

    'Parameters:
    'AppName = Name of Application
    'AppExtension: = File Extension

    'AppCommand = Command Line for Application
    'Example:
    'CreateFileAssociation "Notepad", ".txt", "notepad.exe"

  Dim bAns As Boolean
  Dim sKeyName As String
  Dim sExtName As String
  AppExtension = Trim(AppExtension)
  If Left(AppExtension, 1) <> "." Then Exit Function
  sExtName = Mid(AppExtension, 2) & " File"
  
  bAns = WriteStringToRegistry(HKEY_CLASSES_ROOT, _
    AppExtension, "", sExtName)
    
  If bAns Then bAns = WriteStringToRegistry(HKEY_CLASSES_ROOT, _
      sExtName & "\shell\open\command", "", AppCommand)
      
  CreateFileAssociation = bAns
End Function

Private Function WriteStringToRegistry(hKey As _
  Long, strPath As String, strValue As String, _
  strdata As String) As Boolean
 

Dim bAns As Boolean

On Error GoTo ErrorHandler
   Dim keyhand As Long
   Dim R As Long
   R = RegCreateKey(hKey, strPath, keyhand)
   If R = 0 Then
        R = RegSetValueEx(keyhand, strValue, 0, _
           REG_SZ, ByVal strdata, Len(strdata))
        R = RegCloseKey(keyhand)
    End If
    
   WriteStringToRegistry = (R = 0)

Exit Function

ErrorHandler:
    WriteStringToRegistry = False
    Exit Function
    
End Function

Public Function DeleteOldFiles(DaysOld As Long, _
                               FileSpec As String, _
                               Optional ComparisonDate As Variant) As Boolean

    'PURPOSE: DELETES ALL FILES THAT ARE DaysOld Older than
    'ComparisonDate, which defaults to now

    'RETURNS: True, if succesful
    'False otherwise (e.g., user passes non-date argument
    'deletion fails because file is in use,
    'file doesn't exist, etc.)

    'will not delete readonly, hidden or system files

    Dim sFileSpec As String
    Dim dCompDate As Date
    Dim sFileName As String
    Dim sFileSplit() As String
    Dim iCtr As Integer, iCount As Integer
    Dim sDir As String

    sFileSpec = FileSpec

    If IsMissing(ComparisonDate) Then
        dCompDate = Now
    ElseIf Not IsDate(ComparisonDate) Then
        'client passed wrong type
        DeleteOldFiles = False
        Exit Function
    Else
        dCompDate = CDate(Format(ComparisonDate, "mm/dd/yyyy"))
    End If

    sFileName = Dir(sFileSpec)

    If sFileName = "" Then
        'returns false is file doesn't exist
        DeleteOldFiles = False
        Exit Function
    End If

    Do

        If sFileName = "" Then Exit Do

        If InStr(sFileSpec, "\") > 0 Then
            sFileSplit = Split(sFileSpec, "\")
            iCount = UBound(sFileSplit) - 1

            For iCtr = 0 To iCount
                sDir = sDir & sFileSplit(iCtr) & "\"
            Next

            sFileName = sDir & sFileName
        End If

        On Error GoTo errhandler:

        If DateDiff("d", FileDateTime(sFileName), dCompDate) >= DaysOld Then

            Kill sFileName

        End If

        sFileName = Dir
        sDir = ""
    Loop

    DeleteOldFiles = True

    Exit Function

errhandler:
    DeleteOldFiles = False
    Exit Function
End Function



