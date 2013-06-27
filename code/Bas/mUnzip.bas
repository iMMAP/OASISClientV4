Attribute VB_Name = "mUnzip"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' argv
Private Type UNZIPnames
    s(0 To 1023) As String
End Type

' Callback large "string" (sic)
Private Type CBChar
    ch(0 To 32800) As Byte
End Type

' Callback small "string" (sic)
Private Type CBCh
    ch(0 To 255) As Byte
End Type

' DCL structure
Public Type DCLIST
   ExtractOnlyNewer As Long      ' 1 to extract only newer
   SpaceToUnderScore As Long     ' 1 to convert spaces to underscore
   PromptToOverwrite As Long     ' 1 if overwriting prompts required
   fQuiet As Long                ' 0 = all messages, 1 = few messages, 2 = no messages
   ncflag As Long                ' write to stdout if 1
   ntflag As Long                ' test zip file
   nvflag As Long                ' verbose listing
   nUflag As Long                ' "update" (extract only newer/new files)
   nzflag As Long                ' display zip file comment
   ndflag As Long                ' all args are files/dir to be extracted
   noflag As Long                ' 1 if always overwrite files
   naflag As Long                ' 1 to do end-of-line translation
   nZIflag As Long               ' 1 to get zip info
   C_flag As Long                ' 1 to be case insensitive
   fPrivilege As Long            ' zip file name
   lpszZipFN As String           ' directory to extract to.
   lpszExtractDir As String
End Type

Private Type USERFUNCTION
   ' Callbacks:
   lptrPrnt As Long           ' Pointer to application's print routine
   lptrSound As Long          ' Pointer to application's sound routine.  NULL if app doesn't use sound
   lptrReplace As Long        ' Pointer to application's replace routine.
   lptrPassword As Long       ' Pointer to application's password routine.
   lptrMessage As Long        ' Pointer to application's routine for
                              ' displaying information about specific files in the archive
                              ' used for listing the contents of the archive.
   lptrService As Long        ' callback function designed to be used for allowing the
                              ' app to process Windows messages, or cancelling the operation
                              ' as well as giving option of progress.  If this function returns
                              ' non-zero, it will terminate what it is doing.  It provides the app
                              ' with the name of the archive member it has just processed, as well
                              ' as the original size.
                              
   ' Values filled in after processing:
   lTotalSizeComp As Long     ' Value to be filled in for the compressed total size, excluding
                              ' the archive header and central directory list.
   lTotalSize As Long         ' Total size of all files in the archive
   lCompFactor As Long        ' Overall archive compression factor
   lNumMembers As Long        ' Total number of files in the archive
   cchComment As Integer      ' Flag indicating whether comment in archive.
End Type

Public Type ZIPVERSIONTYPE
   major As Byte
   minor As Byte
   patchlevel As Byte
   not_used As Byte
End Type

Public Type UZPVER
    structlen As Long         ' Length of structure
    flag As Long              ' 0 is beta, 1 uses zlib
    betalevel As String * 10  ' e.g "g BETA"
    date As String * 20       ' e.g. "4 Sep 95" (beta) or "4 September 1995"
    zlib As String * 10       ' e.g. "1.0.5 or NULL"
    Unzip As ZIPVERSIONTYPE
    zipinfo As ZIPVERSIONTYPE
    os2dll As ZIPVERSIONTYPE
    windll As ZIPVERSIONTYPE
End Type

Private Declare Function Wiz_SingleEntryUnzip Lib "vbuzip10.dll" _
  (ByVal ifnc As Long, ByRef ifnv As UNZIPnames, _
   ByVal xfnc As Long, ByRef xfnv As UNZIPnames, _
   dcll As DCLIST, Userf As USERFUNCTION) As Long
Public Declare Sub UzpVersion2 Lib "vbuzip10.dll" (uzpv As UZPVER)

' Object for callbacks:
Private m_cUnzip As cUnzip
Private m_bCancel As Boolean

Private Function plAddressOf(ByVal lPtr As Long) As Long
       ' VB Bug workaround fn
        '<EhHeader>
        On Error GoTo plAddressOf_Err
        '</EhHeader>
100    plAddressOf = lPtr
        '<EhFooter>
        Exit Function

plAddressOf_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.mUnzip.plAddressOf " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub UnzipMessageCallBack( _
      ByVal ucsize As Long, _
      ByVal csiz As Long, _
      ByVal cfactor As Integer, _
      ByVal mo As Integer, _
      ByVal dy As Integer, _
      ByVal yr As Integer, _
      ByVal hh As Integer, _
      ByVal mm As Integer, _
      ByVal c As Byte, _
      ByRef fname As CBCh, _
      ByRef meth As CBCh, _
      ByVal crc As Long, _
      ByVal fCrypt As Byte _
   )
        '<EhHeader>
        On Error GoTo UnzipMessageCallBack_Err
        '</EhHeader>
    Dim sFilename As String
    Dim sFolder As String
    Dim dDate As Date
    Dim sMethod As String
    Dim iPos As Long

       On Error Resume Next
    
       ' Add to unzip class:
100    With m_cUnzip
          ' Parse:
102       sFilename = StrConv(fname.ch, vbUnicode)
104       ParseFileFolder sFilename, sFolder
106       dDate = DateSerial(yr, mo, hh)
108       dDate = dDate + TimeSerial(hh, mm, 0)
110       sMethod = StrConv(meth.ch, vbUnicode)
112       iPos = InStr(sMethod, vbNullChar)
114       If (iPos > 1) Then
116          sMethod = Left$(sMethod, iPos - 1)
          End If
    
118       DebugPrint fCrypt
120       .DirectoryListAddFile sFilename, sFolder, dDate, csiz, crc, ((fCrypt And 64) = 64), cfactor, sMethod
       End With
   
        '<EhFooter>
        Exit Sub

UnzipMessageCallBack_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.mUnzip.UnzipMessageCallBack " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function UnzipPrintCallback( _
      ByRef fname As CBChar, _
      ByVal x As Long _
   ) As Long
        '<EhHeader>
        On Error GoTo UnzipPrintCallback_Err
        '</EhHeader>
    Dim iPos As Long
    Dim sFile As String
       On Error Resume Next
   
       ' Check we've got a message:
100    If x > 1 And x < 1024 Then
          ' If so, then get the readable portion of it:
102       ReDim b(0 To x) As Byte
104       CopyMemory b(0), fname, x
          ' Convert to VB string:
106       sFile = StrConv(b, vbUnicode)
      
          ' Fix up backslashes:
108       ReplaceSection sFile, "/", "\"
      
          ' Tell the caller about it
110       m_cUnzip.ProgressReport sFile
       End If
112    UnzipPrintCallback = 0
        '<EhFooter>
        Exit Function

UnzipPrintCallback_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.mUnzip.UnzipPrintCallback " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function UnzipPasswordCallBack( _
      ByRef pwd As CBCh, _
      ByVal x As Long, _
      ByRef S2 As CBCh, _
      ByRef Name As CBCh _
   ) As Long
        '<EhHeader>
        On Error GoTo UnzipPasswordCallBack_Err
        '</EhHeader>

    Dim bCancel As Boolean
    Dim sPassword As String
    Dim b() As Byte
    Dim lSize As Long

    On Error Resume Next

       ' The default:
100    UnzipPasswordCallBack = 1
    
102    If m_bCancel Then
          Exit Function
       End If
   
       ' Ask for password:
104    m_cUnzip.PasswordRequest sPassword, bCancel
      
106    sPassword = Trim$(sPassword)
   
       ' Cancel out if no useful password:
108    If bCancel Or Len(sPassword) = 0 Then
110       m_bCancel = True
          Exit Function
       End If
   
       ' Put password into return parameter:
112    lSize = Len(sPassword)
114    If lSize > 254 Then
116       lSize = 254
       End If
118    b = StrConv(sPassword, vbFromUnicode)
120    CopyMemory pwd.ch(0), b(0), lSize
   
       ' Ask UnZip to process it:
122    UnzipPasswordCallBack = 0
       
        '<EhFooter>
        Exit Function

UnzipPasswordCallBack_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.mUnzip.UnzipPasswordCallBack " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function UnzipReplaceCallback(ByRef fname As CBChar) As Long
        '<EhHeader>
        On Error GoTo UnzipReplaceCallback_Err
        '</EhHeader>
    Dim eResponse As EUZOverWriteResponse
    Dim iPos As Long
    Dim sFile As String

       On Error Resume Next
100    eResponse = euzDoNotOverwrite
   
       ' Extract the filename:
102    sFile = StrConv(fname.ch, vbUnicode)
104    iPos = InStr(sFile, vbNullChar)
106    If (iPos > 1) Then
108       sFile = Left$(sFile, iPos - 1)
       End If
   
       ' No backslashes:
110    ReplaceSection sFile, "/", "\"
   
       ' Request the overwrite request:
112    m_cUnzip.OverwriteRequest sFile, eResponse
   
       ' Return it to the zipping lib
114    UnzipReplaceCallback = eResponse
   
        '<EhFooter>
        Exit Function

UnzipReplaceCallback_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.mUnzip.UnzipReplaceCallback " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Private Function UnZipServiceCallback(ByRef mname As CBChar, ByVal x As Long) As Long
        '<EhHeader>
        On Error GoTo UnZipServiceCallback_Err
        '</EhHeader>
    Dim iPos As Long
    Dim sInfo As String
    Dim bCancel As Boolean
    
    '-- Always Put This In Callback Routines!
    On Error Resume Next
    
       ' Check we've got a message:
100    If x > 1 And x < 1024 Then
          ' If so, then get the readable portion of it:
102       ReDim b(0 To x) As Byte
104       CopyMemory b(0), mname, x
          ' Convert to VB string:
106       sInfo = StrConv(b, vbUnicode)
108       iPos = InStr(sInfo, vbNullChar)
110       If iPos > 0 Then
112          sInfo = Left$(sInfo, iPos - 1)
          End If
114       ReplaceSection sInfo, "\", "/"
116       m_cUnzip.Service sInfo, bCancel
118       If bCancel Then
120          UnZipServiceCallback = 1
          Else
122          UnZipServiceCallback = 0
          End If
       End If
   
        '<EhFooter>
        Exit Function

UnZipServiceCallback_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.mUnzip.UnZipServiceCallback " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function



Private Sub ParseFileFolder( _
      ByRef sFilename As String, _
      ByRef sFolder As String _
   )
        '<EhHeader>
        On Error GoTo ParseFileFolder_Err
        '</EhHeader>
    Dim iPos As Long
    Dim iLastPos As Long

100    iPos = InStr(sFilename, vbNullChar)
102    If (iPos <> 0) Then
104       sFilename = Left$(sFilename, iPos - 1)
       End If
   
106    iLastPos = ReplaceSection(sFilename, "/", "\")
   
108    If (iLastPos > 1) Then
110       sFolder = Left$(sFilename, iLastPos - 2)
112       sFilename = Mid$(sFilename, iLastPos)
       End If
   
        '<EhFooter>
        Exit Sub

ParseFileFolder_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.mUnzip.ParseFileFolder " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Private Function ReplaceSection(ByRef sString As String, ByVal sToReplace As String, ByVal sReplaceWith As String) As Long
        '<EhHeader>
        On Error GoTo ReplaceSection_Err
        '</EhHeader>
    Dim iPos As Long
    Dim iLastPos As Long
100    iLastPos = 1
       Do
102       iPos = InStr(iLastPos, sString, "/")
104       If (iPos > 1) Then
106          Mid$(sString, iPos, 1) = "\"
108          iLastPos = iPos + 1
          End If
110    Loop While Not (iPos = 0)
112    ReplaceSection = iLastPos

        '<EhFooter>
        Exit Function

ReplaceSection_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.mUnzip.ReplaceSection " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

' Main subroutine
Public Function VBUnzip( _
      cUnzipObject As cUnzip, _
      tDCL As DCLIST, _
      iIncCount As Long, _
      sInc() As String, _
      iExCount As Long, _
      sExc() As String _
   ) As Long
        '<EhHeader>
        On Error GoTo VBUnzip_Err
        '</EhHeader>
    Dim tUser As USERFUNCTION
    Dim lR As Long
    Dim tInc As UNZIPnames
    Dim tExc As UNZIPnames
    Dim i As Long

    On Error GoTo ERRORHANDLER

100    Set m_cUnzip = cUnzipObject
       ' Set Callback addresses
102    tUser.lptrPrnt = plAddressOf(AddressOf UnzipPrintCallback)
104    tUser.lptrSound = 0& ' not supported
106    tUser.lptrReplace = plAddressOf(AddressOf UnzipReplaceCallback)
108    tUser.lptrPassword = plAddressOf(AddressOf UnzipPasswordCallBack)
110    tUser.lptrMessage = plAddressOf(AddressOf UnzipMessageCallBack)
112    tUser.lptrService = plAddressOf(AddressOf UnZipServiceCallback)
        
       ' Set files to include/exclude:
114    If (iIncCount > 0) Then
116       For i = 1 To iIncCount
118          tInc.s(i - 1) = sInc(i)
120       Next i
122       tInc.s(iIncCount) = vbNullChar
       Else
124       tInc.s(0) = vbNullChar
       End If
126    If (iExCount > 0) Then
128       For i = 1 To iExCount
130          tExc.s(i - 1) = sExc(i)
132       Next i
134       tExc.s(iExCount) = vbNullChar
       Else
136       tExc.s(0) = vbNullChar
       End If
138    m_bCancel = False
140    VBUnzip = Wiz_SingleEntryUnzip(iIncCount, tInc, iExCount, tExc, tDCL, tUser)
    
        'DebugPrint "--------------"
        'DebugPrint MYUSER.cchComment
        'DebugPrint MYUSER.TotalSizeComp
        'DebugPrint MYUSER.TotalSize
        'DebugPrint MYUSER.CompFactor
        'DebugPrint MYUSER.NumMembers
        'DebugPrint "--------------"

       Exit Function
   
ERRORHANDLER:
    Dim lErr As Long, sErr As Long
142    lErr = Err.number: sErr = Err.Description
144    VBUnzip = -1
146    Set m_cUnzip = Nothing
148    Err.Raise lErr, App.EXEName & ".VBUnzip", sErr
       Exit Function

        '<EhFooter>
        Exit Function

VBUnzip_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.mUnzip.VBUnzip " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
