VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUnzip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EUZMsgLevel
   euzAllMessages = 0
   euzPartialMessages = 1
   euzNoMessages = 2
End Enum
Public Enum EUZOverWriteResponse
   euzDoNotOverwrite = 100
   euzOverwriteThisFile = 102
   euzOverwriteAllFiles = 103
   euzOverwriteNone = 104
End Enum

Private m_sZipFile As String
Private m_sUnzipFolder As String
Private Type tZipContents
   sName As String
   sFolder As String
   lSize As Long
   lPackedSize As Long
   lFactor As Long
   sMethod As String
   dDate As Date
   lCrc As Long
   fEncryped As Boolean
   fSelected As Boolean
End Type
Private m_tZipContents() As tZipContents
Private m_iCount As Long
Private m_tDCL As DCLIST

Public Event Cancel(ByVal sMsg As String, ByRef bCancel As Boolean)
Public Event OverWritePrompt(ByVal sFile As String, ByRef eResponse As EUZOverWriteResponse)
Public Event PasswordRequest(ByRef sPassword As String, ByRef bCancel As Boolean)
Public Event Progress(ByVal lCount As Long, ByVal sMsg As String)

Public Property Get ExtractOnlyNewer() As Boolean
        '<EhHeader>
        On Error GoTo ExtractOnlyNewer_Err
        '</EhHeader>
100    ExtractOnlyNewer = (m_tDCL.ExtractOnlyNewer <> 0)      ' 1=extract only newer
        '<EhFooter>
        Exit Property

ExtractOnlyNewer_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.ExtractOnlyNewer", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Let ExtractOnlyNewer(ByVal bState As Boolean)
        '<EhHeader>
        On Error GoTo ExtractOnlyNewer_Err
        '</EhHeader>
100    m_tDCL.ExtractOnlyNewer = Abs(bState)      ' 1=extract only newer
        '<EhFooter>
        Exit Property

ExtractOnlyNewer_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.ExtractOnlyNewer", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get SpaceToUnderScore() As Boolean
        '<EhHeader>
        On Error GoTo SpaceToUnderScore_Err
        '</EhHeader>
100    SpaceToUnderScore = (m_tDCL.SpaceToUnderScore <> 0)    ' 1=convert space to underscore
        '<EhFooter>
        Exit Property

SpaceToUnderScore_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.SpaceToUnderScore", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Let SpaceToUnderScore(ByVal bState As Boolean)
        '<EhHeader>
        On Error GoTo SpaceToUnderScore_Err
        '</EhHeader>
100    m_tDCL.SpaceToUnderScore = Abs(bState)
        '<EhFooter>
        Exit Property

SpaceToUnderScore_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.SpaceToUnderScore", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get PromptToOverwrite() As Boolean
        '<EhHeader>
        On Error GoTo PromptToOverwrite_Err
        '</EhHeader>
100    PromptToOverwrite = (m_tDCL.PromptToOverwrite <> 0)
        '<EhFooter>
        Exit Property

PromptToOverwrite_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.PromptToOverwrite", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Let PromptToOverwrite(ByVal bState As Boolean)
        '<EhHeader>
        On Error GoTo PromptToOverwrite_Err
        '</EhHeader>
100    m_tDCL.PromptToOverwrite = Abs(bState)
        '<EhFooter>
        Exit Property

PromptToOverwrite_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.PromptToOverwrite", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get MessageLevel() As EUZMsgLevel
        '<EhHeader>
        On Error GoTo MessageLevel_Err
        '</EhHeader>
100    MessageLevel = m_tDCL.fQuiet
        '<EhFooter>
        Exit Property

MessageLevel_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.MessageLevel", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Let MessageLevel(ByVal eLevel As EUZMsgLevel)
        '<EhHeader>
        On Error GoTo MessageLevel_Err
        '</EhHeader>
100    m_tDCL.fQuiet = eLevel
        '<EhFooter>
        Exit Property

MessageLevel_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.MessageLevel", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get TestZip() As Boolean
        '<EhHeader>
        On Error GoTo TestZip_Err
        '</EhHeader>
100    TestZip = (m_tDCL.ntflag <> 0)
        '<EhFooter>
        Exit Property

TestZip_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.TestZip", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Let TestZip(ByVal bState As Boolean)
        '<EhHeader>
        On Error GoTo TestZip_Err
        '</EhHeader>
100    m_tDCL.ntflag = Abs(bState)
        '<EhFooter>
        Exit Property

TestZip_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.TestZip", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get UseFolderNames() As Boolean
        '<EhHeader>
        On Error GoTo UseFolderNames_Err
        '</EhHeader>
100    UseFolderNames = (m_tDCL.ndflag <> 0)
        '<EhFooter>
        Exit Property

UseFolderNames_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.UseFolderNames", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Let UseFolderNames(ByVal bState As Boolean)
        '<EhHeader>
        On Error GoTo UseFolderNames_Err
        '</EhHeader>
100    m_tDCL.ndflag = Abs(bState)
        '<EhFooter>
        Exit Property

UseFolderNames_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.UseFolderNames", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get OverwriteExisting() As Boolean
        '<EhHeader>
        On Error GoTo OverwriteExisting_Err
        '</EhHeader>
100    OverwriteExisting = (m_tDCL.noflag <> 0)
        '<EhFooter>
        Exit Property

OverwriteExisting_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.OverwriteExisting", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Let OverwriteExisting(ByVal bState As Boolean)
        '<EhHeader>
        On Error GoTo OverwriteExisting_Err
        '</EhHeader>
100    m_tDCL.noflag = Abs(bState)
        '<EhFooter>
        Exit Property

OverwriteExisting_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.OverwriteExisting", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get ConvertCRToCRLF() As Boolean
        '<EhHeader>
        On Error GoTo ConvertCRToCRLF_Err
        '</EhHeader>
100    ConvertCRToCRLF = (m_tDCL.naflag <> 0)
        '<EhFooter>
        Exit Property

ConvertCRToCRLF_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.ConvertCRToCRLF", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Let ConvertCRToCRLF(ByVal bState As Boolean)
        '<EhHeader>
        On Error GoTo ConvertCRToCRLF_Err
        '</EhHeader>
100    m_tDCL.naflag = Abs(bState)
        '<EhFooter>
        Exit Property

ConvertCRToCRLF_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.ConvertCRToCRLF", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get CaseSensitiveFileNames() As Boolean
        '<EhHeader>
        On Error GoTo CaseSensitiveFileNames_Err
        '</EhHeader>
100    CaseSensitiveFileNames = m_tDCL.C_flag
        '<EhFooter>
        Exit Property

CaseSensitiveFileNames_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.CaseSensitiveFileNames", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Let CaseSensitiveFileNames(ByVal bState As Boolean)
        '<EhHeader>
        On Error GoTo CaseSensitiveFileNames_Err
        '</EhHeader>
100    m_tDCL.C_flag = Abs(bState)
        '<EhFooter>
        Exit Property

CaseSensitiveFileNames_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.CaseSensitiveFileNames", _
                  "cUnzip component failure"
        '</EhFooter>
End Property

Friend Sub DirectoryListAddFile( _
      ByVal sFilename As String, _
      ByVal sFolder As String, _
      ByVal dDate As Date, _
      ByVal lSize As Long, _
      ByVal lCrc As Long, _
      ByVal fEncrypted As Boolean, _
      ByVal lFactor As Long, _
      ByVal sMethod As String _
   )
        '<EhHeader>
        On Error GoTo DirectoryListAddFile_Err
        '</EhHeader>
100    If (sFilename <> vbNullChar) And Len(sFilename) > 0 Then
102       m_iCount = m_iCount + 1
104       ReDim Preserve m_tZipContents(1 To m_iCount) As tZipContents
106       With m_tZipContents(m_iCount)
108          .sName = sFilename
110          .sFolder = sFolder
112          .dDate = dDate
114          .lSize = lSize
116          .lCrc = lCrc
118          .lFactor = lFactor
120          .sMethod = sMethod
122          .fEncryped = fEncrypted
             ' Default to selected:
124          .fSelected = True
          End With
       End If
        '<EhFooter>
        Exit Sub

DirectoryListAddFile_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.DirectoryListAddFile", _
                  "cUnzip component failure"
        '</EhFooter>
End Sub
Friend Sub OverwriteRequest( _
      ByVal sFile As String, _
      ByRef eResponse As EUZOverWriteResponse _
   )
        '<EhHeader>
        On Error GoTo OverwriteRequest_Err
        '</EhHeader>
100    RaiseEvent OverWritePrompt(sFile, eResponse)
        '<EhFooter>
        Exit Sub

OverwriteRequest_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.OverwriteRequest", _
                  "cUnzip component failure"
        '</EhFooter>
End Sub
Friend Sub ProgressReport( _
      ByVal sMsg As String _
   )
        '<EhHeader>
        On Error GoTo ProgressReport_Err
        '</EhHeader>
100    RaiseEvent Progress(1, sMsg)
        '<EhFooter>
        Exit Sub

ProgressReport_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.ProgressReport", _
                  "cUnzip component failure"
        '</EhFooter>
End Sub
Friend Sub Service( _
      ByVal sMsg As String, _
      ByRef bCancel As Boolean _
   )
        '<EhHeader>
        On Error GoTo Service_Err
        '</EhHeader>
100    RaiseEvent Cancel(sMsg, bCancel)
        '<EhFooter>
        Exit Sub

Service_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.Service", _
                  "cUnzip component failure"
        '</EhFooter>
End Sub
      
Friend Sub PasswordRequest( _
      ByRef sPassword As String, _
      ByRef bCancel As Boolean _
   )
        '<EhHeader>
        On Error GoTo PasswordRequest_Err
        '</EhHeader>
100    RaiseEvent PasswordRequest(sPassword, bCancel)
        '<EhFooter>
        Exit Sub

PasswordRequest_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.PasswordRequest", _
                  "cUnzip component failure"
        '</EhFooter>
End Sub
Public Property Get FileCount() As Long
        '<EhHeader>
        On Error GoTo FileCount_Err
        '</EhHeader>
100    FileCount = m_iCount
        '<EhFooter>
        Exit Property

FileCount_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.FileCount", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get Filename(ByVal lIndex As Long) As String
        '<EhHeader>
        On Error GoTo Filename_Err
        '</EhHeader>
100    Filename = m_tZipContents(lIndex).sName
        '<EhFooter>
        Exit Property

Filename_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.Filename", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get FileDirectory(ByVal lIndex As Long) As String
        '<EhHeader>
        On Error GoTo FileDirectory_Err
        '</EhHeader>
100    FileDirectory = m_tZipContents(lIndex).sFolder
        '<EhFooter>
        Exit Property

FileDirectory_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.FileDirectory", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get FileSize(ByVal lIndex As Long) As Long
        '<EhHeader>
        On Error GoTo FileSize_Err
        '</EhHeader>
100    FileSize = m_tZipContents(lIndex).lSize
        '<EhFooter>
        Exit Property

FileSize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.FileSize", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get FilePackedSize(ByVal lIndex As Long) As Long
        '<EhHeader>
        On Error GoTo FilePackedSize_Err
        '</EhHeader>
100    FilePackedSize = m_tZipContents(lIndex).lSize * m_tZipContents(lIndex).lFactor \ 100
        '<EhFooter>
        Exit Property

FilePackedSize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.FilePackedSize", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get FileCompressionRatio(ByVal lIndex As Long) As Long
        '<EhHeader>
        On Error GoTo FileCompressionRatio_Err
        '</EhHeader>
100    FileCompressionRatio = m_tZipContents(lIndex).lFactor
        '<EhFooter>
        Exit Property

FileCompressionRatio_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.FileCompressionRatio", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get FileDate(ByVal lIndex As Long) As Date
        '<EhHeader>
        On Error GoTo FileDate_Err
        '</EhHeader>
100    FileDate = m_tZipContents(lIndex).dDate
        '<EhFooter>
        Exit Property

FileDate_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.FileDate", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get FileCRC(ByVal lIndex As Long) As Long
        '<EhHeader>
        On Error GoTo FileCRC_Err
        '</EhHeader>
100    FileCRC = m_tZipContents(lIndex).lCrc
        '<EhFooter>
        Exit Property

FileCRC_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.FileCRC", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get FileCompressionMethod(ByVal lIndex As Long) As String
        '<EhHeader>
        On Error GoTo FileCompressionMethod_Err
        '</EhHeader>
100    FileCompressionMethod = m_tZipContents(lIndex).sMethod
        '<EhFooter>
        Exit Property

FileCompressionMethod_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.FileCompressionMethod", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get FileEncrypted(ByVal lIndex As Long) As Boolean
        '<EhHeader>
        On Error GoTo FileEncrypted_Err
        '</EhHeader>
100    FileEncrypted = m_tZipContents(lIndex).fEncryped
        '<EhFooter>
        Exit Property

FileEncrypted_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.FileEncrypted", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get FileSelected(ByVal lIndex As Long) As Boolean
        '<EhHeader>
        On Error GoTo FileSelected_Err
        '</EhHeader>
100    FileSelected = m_tZipContents(lIndex).fSelected
        '<EhFooter>
        Exit Property

FileSelected_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.FileSelected", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Let FileSelected(ByVal lIndex As Long, ByVal bState As Boolean)
        '<EhHeader>
        On Error GoTo FileSelected_Err
        '</EhHeader>
100    m_tZipContents(lIndex).fSelected = bState
        '<EhFooter>
        Exit Property

FileSelected_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.FileSelected", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Function Directory() As Long
        '<EhHeader>
        On Error GoTo Directory_Err
        '</EhHeader>
    Dim s(0 To 0) As String
100    m_tDCL.lpszZipFN = m_sZipFile
102    m_tDCL.lpszExtractDir = vbNullChar
104    m_tDCL.nvflag = 1
106    mUnzip.VBUnzip Me, m_tDCL, 0, s(), 0, s()
        '<EhFooter>
        Exit Function

Directory_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.Directory", _
                  "cUnzip component failure"
        '</EhFooter>
End Function
Public Function Unzip() As Boolean
        '<EhHeader>
        On Error GoTo Unzip_Err
        '</EhHeader>
    Dim sInc() As String
    Dim iIncCount As Long
    Dim s() As String
    Dim i As Long
100    If (m_sZipFile <> "") Then
102       If (m_iCount > 0) Then
104          For i = 1 To m_iCount
106             If (m_tZipContents(i).fSelected) Then
108                iIncCount = iIncCount + 1
110                ReDim Preserve sInc(1 To iIncCount) As String
112                sInc(iIncCount) = ReverseSlashes(m_tZipContents(i).sFolder, m_tZipContents(i).sName)
                End If
114          Next i
116          If (iIncCount = m_iCount) Then
118             iIncCount = 0
120             ReDim sInc(0 To 0)
             End If
          End If
122       m_tDCL.lpszZipFN = m_sZipFile
124       m_tDCL.nvflag = 0
126       m_tDCL.lpszExtractDir = m_sUnzipFolder
128       Unzip = (mUnzip.VBUnzip(Me, m_tDCL, iIncCount, sInc(), 0, s()) <> 0)
       End If
        '<EhFooter>
        Exit Function

Unzip_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.Unzip", _
                  "cUnzip component failure"
        '</EhFooter>
End Function
Private Function ReverseSlashes( _
      ByVal sFolder As String, _
      ByVal sFile As String _
   ) As String
        '<EhHeader>
        On Error GoTo ReverseSlashes_Err
        '</EhHeader>
    Dim sOut As String
    Dim iPos As Long, iLastPos As Long

100    If Len(sFolder) > 0 And sFolder <> vbNullChar Then
102       sOut = sFolder & "/" & sFile
104       iLastPos = 1
          Do
106          iPos = InStr(iLastPos, sOut, "\")
108          If (iPos <> 0) Then
110             Mid$(sOut, iPos, 1) = "/"
112             iLastPos = iPos + 1
             End If
114       Loop While iPos <> 0
116       ReverseSlashes = sOut
       Else
118       ReverseSlashes = sFile
       End If
        '<EhFooter>
        Exit Function

ReverseSlashes_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.ReverseSlashes", _
                  "cUnzip component failure"
        '</EhFooter>
End Function
Public Property Let UnzipFolder(ByVal sFolder As String)
        '<EhHeader>
        On Error GoTo UnzipFolder_Err
        '</EhHeader>
100    m_sUnzipFolder = sFolder
        '<EhFooter>
        Exit Property

UnzipFolder_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.UnzipFolder", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get UnzipFolder() As String
        '<EhHeader>
        On Error GoTo UnzipFolder_Err
        '</EhHeader>
100    UnzipFolder = m_sUnzipFolder
102    m_tDCL.lpszExtractDir = m_sUnzipFolder
        '<EhFooter>
        Exit Property

UnzipFolder_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.UnzipFolder", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Property Get ZipFile() As String
        '<EhHeader>
        On Error GoTo ZipFile_Err
        '</EhHeader>
100    ZipFile = m_sZipFile
        '<EhFooter>
        Exit Property

ZipFile_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.ZipFile", _
                  "cUnzip component failure"
        '</EhFooter>
End Property
Public Sub GetVersion( _
      ByRef lMajor As Long, _
      ByRef lMinor As Long, _
      ByRef lRevision As Long, _
      ByRef dDate As Date, _
      ByRef sExtraInfo As String _
   )
        '<EhHeader>
        On Error GoTo GetVersion_Err
        '</EhHeader>
    Dim tVer As UZPVER
    Dim iPos As Long

       ' Set Version space
100    With tVer
102       .structlen = Len(tVer)
104       .betalevel = Space$(9) & vbNullChar
106       .date = Space$(19) & vbNullChar
108       .zlib = Space$(9) & vbNullChar
       End With
    
       ' Get version
110    UzpVersion2 tVer
112    iPos = InStr(tVer.betalevel, vbNullChar)
114    If (iPos > 1) Then
116       tVer.betalevel = Left$(tVer.betalevel, iPos - 1)
       End If
118    sExtraInfo = tVer.betalevel
  
       ' Date..
120    DebugPrint tVer.date

122    lMajor = tVer.windll.major
124    lMinor = tVer.windll.minor
126    lRevision = tVer.windll.patchlevel
      
        '<EhFooter>
        Exit Sub

GetVersion_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.GetVersion", _
                  "cUnzip component failure"
        '</EhFooter>
End Sub

Public Property Let ZipFile(ByVal sFile As String)
        '<EhHeader>
        On Error GoTo ZipFile_Err
        '</EhHeader>
100    m_sZipFile = sFile
102    m_iCount = 0
104    Erase m_tZipContents
        '<EhFooter>
        Exit Property

ZipFile_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.cUnzip.ZipFile", _
                  "cUnzip component failure"
        '</EhFooter>
End Property


