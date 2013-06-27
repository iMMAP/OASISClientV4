VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clIniReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If Win32 Then
    Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
    Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
#End If

Private mPath           As String
Private mSection        As String
Private mKey            As String
Private mValue          As String
Private mDefault        As String

Public Property Let Path(cPath As String)
    mPath = cPath
End Property

Public Property Let Section(cSection As String)
    mSection = cSection
End Property

Public Property Let Key(cKey As String)
    mKey = cKey
End Property

Public Property Let Value(cValue As String)
    mValue = cValue
End Property

Public Property Let Default(cDefault As String)
    mDefault = cDefault
End Property

Public Property Get Value() As String
Dim sBuf As String
Dim iSize As String
Dim iRetCode As Integer
    ErrH1
    sBuf = Space$(255)
    iSize = Len(sBuf)
    iRetCode = GetPrivateProfileString(mSection, mKey, mDefault, sBuf, iSize, mPath)
    If (iSize > 0) Then
        Value = Left$(sBuf, iRetCode)
    Else
        Value = ""
    End If

End Property

Function ErrH1() As Boolean
    If mPath <> "" Then
        If mSection <> "" Then
            If mKey <> "" Then
               ErrH1 = False
               Exit Function
               ' If mValue <> "" Then
               '     ErrH1 = False
               ' Else
               '     MsgBox "Value not set!", vbExclamation
               '     ErrH1 = True
               '     Exit Function
               ' End If
            Else
                MsgBox "Key not set!", vbExclamation
                ErrH1 = True
                Exit Function
            End If
        Else
            'MsgBox "Section not set!", vbExclamation
            ErrH1 = True
            Exit Function
        End If
    Else
       MsgBox "Path not set!", vbExclamation
       ErrH1 = True
       Exit Function
    End If
End Function

Public Function GetAll() As String()
    Dim str1 As String
    Dim int1 As Integer
    Dim FileCell As Integer
    Dim str2(255) As String
    FileCell = FreeFile
    
    Open mPath For Input As FileCell
    
    Do Until EOF(FileCell)
        Line Input #FileCell, str1
        
        str2(int1) = str1
        int1 = int1 + 1
    Loop
    GetAll = str2
    Close FileCell
End Function

Public Sub DeleteKey()
    If mSection = "" Or mKey = "" Then Exit Sub
    Call WritePrivateProfileString(mSection, mKey, 0&, mPath)
    'MsgBox "A key deleted!", vbInformation
End Sub
Public Sub DeleteSection()
    If mSection = "" Then Exit Sub
    Call WritePrivateProfileString(mSection, 0&, 0&, mPath)
    'MsgBox "Section deleted!", vbInformation
End Sub

Public Sub AddKeyWithValue()
    If ErrH1 = True Then Exit Sub
    Call WritePrivateProfileString(mSection, mKey, mValue, mPath)
    'MsgBox "Key saved!", vbInformation
End Sub
Public Sub AddNewSection()
    If ErrH1 = True Then Exit Sub
    Call WritePrivateProfileString(mSection, mKey, mValue, mPath)
    'MsgBox "New section saved!", vbInformation
End Sub


