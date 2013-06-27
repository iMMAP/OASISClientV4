VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************************************************************************************************************
'*  Author: Nabhan Ahmed                                                                                    *
'*  Modified By: Donathan Hutchings                                                                         *
'*  Date: 9/9/2003                                                                                          *
'*                                                                                                          *
'*  Description:    Basically, this program shows you how to copy files byte by byte.                       *
'*                  Everytime it reads 4 kbs from the source file and write them in the                     *
'*                  destination file until it reads all bytes in the srouce file. There                     *
'*                  is a bar that shows the copying progress, and there is a label that                     *
'*                  shows the copying percentage. In the new update, you can use the browse                 *
'*                  button to find the source file you want to copy. I'm sure you'll find                   *
'*                  the code very easy to understand, and that you'll like it.( Would you                   *
'*                  please vote for this code )                                                             *
'*                                                                                                          *
'*  Changes:        The FileCopy routine was transfered to a self contained class.  It checks               *
'*                  for proper input for the FileCopy method to work.  It is now completey                  *
'*                  portable.  All credit should go to Nabhan Ahmed for a great routine.                    *
'*                  Vote for him at:                                                                        *
'*                  http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=48319&lngWId=1      *
'*                                                                                                          *
'************************************************************************************************************

Option Explicit

Private m_SourceFile    As String
Private m_TargetFile    As String
Private m_ByteSize      As Double
Private m_Verify        As Boolean

Event Progress(Value As Integer, BytesCopied As Double)

Public Property Get SourceFile() As String
    SourceFile = m_SourceFile
End Property

Public Property Let SourceFile(sNewValue As String)
    m_SourceFile = sNewValue
End Property

Public Property Get TargetFile() As String
    TargetFile = m_TargetFile
End Property

Public Property Let TargetFile(sNewValue As String)
    m_TargetFile = sNewValue
End Property

'How many bites to copy at a time.
Public Property Get ByteSize() As Double
    ByteSize = m_ByteSize
End Property

Public Property Let ByteSize(sNewValue As Double)
    m_ByteSize = sNewValue
End Property

Public Property Get Verify() As Boolean
    Verify = m_Verify
End Property

Public Property Let Verify(bNewValue As Boolean)
    m_Verify = bNewValue
End Property

Public Function CopyFile(Optional ErrCode As Long = 0, Optional ErrMsg As String) As Boolean
    Dim SrcFileLen As Long
    Dim nSF, nDF As Integer
    Dim Chunk As String
    Dim BytesCopied As Double
    Dim TargetPath As String
    
    On Local Error GoTo ReportError
    
    Screen.MousePointer = vbHourglass
    
    'Check for correct values
    Select Case True
    Case Len(Trim(m_SourceFile)) = 0
        ErrCode = 100000
        ErrMsg = "Please enter a source file..."
        GoTo ReportError
    Case Len(Trim(m_TargetFile)) = 0
        ErrCode = 100010
        ErrMsg = "Please enter a target file..."
        GoTo ReportError
    Case Not FileExists(m_SourceFile)
        ErrCode = 100020
        ErrMsg = "The source file does not exist..."
        GoTo ReportError
    Case FileExists(m_TargetFile)
        If Not MsgBox(m_TargetFile & " already exists.  Overwrite?", _
                      vbQuestion + vbApplicationModal + vbYesNo) = vbYes Then
            ErrCode = 100030
            ErrMsg = "File copy canceled..."
            GoTo ReportError
        End If
    End Select
    
    'If target path does not exist, then create it.
    TargetPath = GetPathPart(m_TargetFile)
    If Not DirExists(TargetPath) Then
        CreateDirectoryStruct TargetPath
    End If
    
    'Verify ByteSize
    If m_ByteSize <= 0 Then m_ByteSize = 4096   '4kb
    
    'Get source file length
    SrcFileLen = FileLen(m_SourceFile)
    
    'Open both files
    nSF = FreeFile
    Open m_SourceFile For Binary As nSF
    nDF = FreeFile
    Open m_TargetFile For Binary As nDF
    
    'How many bytes to get each time
    BytesCopied = 0
    
    'Keep copying until finishing all bytes
    Do While BytesCopied < SrcFileLen
        'Check how many bytes left
        If m_ByteSize < (SrcFileLen - BytesCopied) Then
            'Copy 4 KBytes
            Chunk = Space(m_ByteSize)
            Get #nSF, , Chunk
        Else
            'Copy the rest
            Chunk = Space(SrcFileLen - BytesCopied)
            Get #nSF, , Chunk
        End If
        
        BytesCopied = BytesCopied + Len(Chunk)
        
        'Show Percentage
        RaiseEvent Progress(Int(BytesCopied / SrcFileLen * 100), BytesCopied)
        
        'Put data in destination file
        Put #nDF, , Chunk
    Loop
    
    'Hide progress bar
    RaiseEvent Progress(0, BytesCopied)
    
    'Optional Verify
    If m_Verify Then
        If Not VerifyFile(ErrCode, ErrMsg) Then
            Call WipeFileClean(m_TargetFile, ErrCode, ErrMsg)
            GoTo ReportError
        End If
    End If
    
    'Close files
    Close #nSF
    Close #nDF

    Screen.MousePointer = vbDefault
    CopyFile = True
    Exit Function
    
ReportError:
    CopyFile = False
    'Close files
    Close #nSF
    Close #nDF
    'Hide progress bar
    RaiseEvent Progress(0, BytesCopied)
    Screen.MousePointer = vbDefault
    If ErrCode = 0 Then
        ErrCode = Err.Number
        ErrMsg = Err.Description
    End If
    Exit Function
    Resume 0
End Function

Private Function FileExists(Filename As String) As Integer
    Dim i As Integer
    
    On Local Error Resume Next
    i = Len(Dir$(Filename$))
    If Err Or i = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If
    On Local Error GoTo 0
End Function

Private Sub CreateDirectoryStruct(CreateThisPath As String)
    'do initial check
    Dim ret As Boolean, Temp$, ComputerName As String, IntoItCount As Integer, x%, WakeString As String
    Dim MadeIt As Integer
    If Dir$(CreateThisPath, vbDirectory) <> "" Then Exit Sub
    'is this a network path?


    If Left$(CreateThisPath, 2) = "\\" Then ' this is a UNC NetworkPath
        'must extract the machine name first, th
        '     en get to the first folder
        IntoItCount = 3
        ComputerName = Mid$(CreateThisPath, IntoItCount, InStr(IntoItCount, CreateThisPath, "\") - IntoItCount)
        IntoItCount = IntoItCount + Len(ComputerName) + 1
        IntoItCount = InStr(IntoItCount, CreateThisPath, "\") + 1
        'temp = Mid$(CreateThisPath, IntoItCount
        '     , x)
    Else ' this is a regular path
        IntoItCount = 4
    End If
    WakeString = Left$(CreateThisPath, IntoItCount - 1)
    'start a loop through the CreateThisPath
    '     string


    Do
        x = InStr(IntoItCount, CreateThisPath, "\")


        If x <> 0 Then
            x = x - IntoItCount
            Temp = Mid$(CreateThisPath, IntoItCount, x)
        Else
            Temp = Mid$(CreateThisPath, IntoItCount)
        End If
        IntoItCount = IntoItCount + Len(Temp) + 1
        Temp = WakeString + Temp
        'Create a directory if it doesn't alread
        '     y exist
        ret = (Dir$(Temp, vbDirectory) <> "")


        If Not ret Then
            'ret& = CreateDirectory(temp, Security)
            MkDir Temp
        End If
        IntoItCount = IntoItCount 'track where we are in the String
        WakeString = Left$(CreateThisPath, IntoItCount - 1)
    Loop While WakeString <> CreateThisPath
End Sub

Private Function DirExists(strDir As String) As Boolean
  
  On Error GoTo PROC_ERR

  DirExists = Len(Dir$(strDir & "\.", vbDirectory)) > 0
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Resume PROC_EXIT
  
End Function

Public Function VerifyFile(Optional ErrCode As Long = 0, Optional ErrMsg As String) As Boolean
    Dim SrcFileLen As Long
    Dim nSF, nDF As Integer
    Dim Chunk(1) As String
    Dim BytesVerified As Double

    On Local Error GoTo ReportError
    
    Screen.MousePointer = vbHourglass
    
    'Check for correct values
    Select Case True
    Case Len(Trim(m_SourceFile)) = 0
        ErrCode = 100000
        ErrMsg = "Please enter a source file..."
        GoTo ReportError
    Case Len(Trim(m_TargetFile)) = 0
        ErrCode = 100010
        ErrMsg = "Please enter a target file..."
        GoTo ReportError
    Case Not FileExists(m_SourceFile)
        ErrCode = 100020
        ErrMsg = "The source file does not exist..."
        GoTo ReportError
    End Select
    
    'Verify ByteSize
    If m_ByteSize <= 0 Then m_ByteSize = 4096   '4kb
    
    'Get source file length
    SrcFileLen = FileLen(m_SourceFile)
    
    'Open both files
    nSF = FreeFile
    Open m_SourceFile For Binary As nSF
    nDF = FreeFile
    Open m_TargetFile For Binary As nDF
    
    'How many bytes to get each time
    BytesVerified = 0
    
    'Keep copying until finishing all bytes
    Do While BytesVerified < SrcFileLen
        'Check how many bytes left
        If m_ByteSize < (SrcFileLen - BytesVerified) Then
            'Copy 4 KBytes
            Chunk(0) = Space(m_ByteSize)
            Chunk(1) = Space(m_ByteSize)
            Get #nSF, , Chunk(0)
            Get #nDF, , Chunk(1)
        Else
            'Copy the rest
            Chunk(0) = Space(SrcFileLen - BytesVerified)
            Chunk(1) = Space(SrcFileLen - BytesVerified)
            Get #nSF, , Chunk(0)
            Get #nDF, , Chunk(1)
        End If
        
        'Compare the two chunks.
        If InStr(1, Chunk(0), Chunk(1), vbBinaryCompare) = 0 Then
            ErrCode = 110000
            ErrMsg = "Target file is not the same as the source file..."
            GoTo ReportError
        End If
        
        BytesVerified = BytesVerified + Len(Chunk(0))
        
        'Show Percentage
        RaiseEvent Progress(Int(BytesVerified / SrcFileLen * 100), BytesVerified)
        
        'Put data in destination file
    Loop
    
    'Hide progress bar
    RaiseEvent Progress(0, BytesVerified)
    
    'Close files
    Close #nSF
    Close #nDF

    Screen.MousePointer = vbDefault
    VerifyFile = True
    Exit Function
    
ReportError:
    VerifyFile = False
    'Close files
    Close #nSF
    Close #nDF
    'Hide progress bar
    RaiseEvent Progress(0, BytesVerified)
    Screen.MousePointer = vbDefault
    If ErrCode = 0 Then
        ErrCode = Err.Number
        ErrMsg = Err.Description
    End If
    Exit Function
    Resume 0
End Function

Private Function WipeFileClean(sFileName As String, ErrCode As Long, ErrMsg As String) As Boolean
    Dim Block1 As String
    Dim Block2 As String
    Dim Blocks As Long
    Dim Block3 As String
    Dim Block4 As String
    Dim Repetition As Integer
    Dim hFileHandle As Integer
    Dim iLoop As Long
    Dim Offset As Long
    Dim LenOfFile As Long
    
    Const BlockSize = 4096
    
    On Local Error GoTo ReportError
    
    LenOfFile = FileLen(sFileName)
    Block1 = String(BlockSize, "X")
    Block2 = String(BlockSize, " ")
    Block3 = String(BlockSize, "1")
    Block4 = String(BlockSize, "0")
    hFileHandle = FreeFile
    Open sFileName For Binary As hFileHandle
        Blocks = (LOF(hFileHandle) \ BlockSize) + 1
        For iLoop = 1 To Blocks
          Offset = Seek(hFileHandle)
          Put hFileHandle, , Block1
          Put hFileHandle, Offset, Block2
          Put hFileHandle, Offset, Block3
          Put hFileHandle, Offset, Block4
        Next iLoop
    Close hFileHandle
    Kill sFileName
    WipeFileClean = True
    Exit Function
    
ReportError:
    WipeFileClean = False
    If ErrCode = 0 Then
        ErrCode = Err.Number
        ErrMsg = Err.Description
    End If
    Exit Function
End Function

Private Function GetPathPart(strPath As String) As String
  Dim intCounter As Integer

  On Error GoTo PROC_ERR
  
  ' Parse the string backwards
  For intCounter = Len(strPath) To 1 Step -1
    ' Short-circuit when we reach the slash
    If Mid$(strPath, intCounter, 1) = "\" Then
      Exit For
    End If
  Next intCounter

  ' Return the value
  GetPathPart = Left$(strPath, intCounter)

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  Resume PROC_EXIT
  
End Function

