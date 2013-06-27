Attribute VB_Name = "modComDialog"
Option Explicit

Const MAX_PATH As Long = 260

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

Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_EXPLORER = &H80000
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000

Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Private Declare Function GetOpenFileName _
                Lib "comdlg32.dll" _
                Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName _
                Lib "comdlg32.dll" _
                Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Function FILE_DIALOG(frmForm As Form, _
                            bSaveDialog As Boolean, _
                            ByVal stitle As String, _
                            ByVal sFilter As String, _
                            Optional ByVal sFileName As String, _
                            Optional ByVal sExtention As String, _
                            Optional ByVal sInitDir As String) As String
        '<EhHeader>
        On Error GoTo FILE_DIALOG_Err
        '</EhHeader>
        Dim OFN As OPENFILENAME, lReturn As Long
100     frmForm.Enabled = False
102     sFileName = sFileName + String(MAX_PATH - Len(sFileName), 0)

104     With OFN
106         .lStructSize = Len(OFN)
108         .hWndOwner = frmForm.hwnd
110         .hInstance = App.hInstance
112         .lpstrFilter = Replace(sFilter, "|", Chr$(0))
114         .lpstrFile = sFileName
116         .nMaxFile = MAX_PATH
118         .lpstrFileTitle = Space$(MAX_PATH - 1)
120         .nMaxFileTitle = MAX_PATH
122         .lpstrInitialDir = sInitDir
124         .lpstrTitle = stitle
126         .Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
128         .lpstrDefExt = sExtention
        End With

130     If bSaveDialog Then lReturn = GetSaveFileName(OFN) Else lReturn = GetOpenFileName(OFN)
132     If lReturn <> 0 Then FILE_DIALOG = Left$(OFN.lpstrFile + vbNullChar, InStr(1, OFN.lpstrFile + vbNullChar, vbNullChar) - 1)
134     frmForm.Enabled = True
        '<EhFooter>
        Exit Function

FILE_DIALOG_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modComDialog.FILE_DIALOG " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub RID_FILE(ByVal sFileName As String)
        '<EhHeader>
        On Error GoTo RID_FILE_Err
        '</EhHeader>

100     If FILE_EXISTS(sFileName) Then
102         SetAttr sFileName, vbNormal
104         Kill sFileName
        End If

        '<EhFooter>
        Exit Sub

RID_FILE_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modComDialog.RID_FILE " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function FILE_TITLE_ONLY(sFileName As String, _
                                Optional bReturnDirectory As Boolean) As String
        '<EhHeader>
        On Error GoTo FILE_TITLE_ONLY_Err
        '</EhHeader>
100     FILE_TITLE_ONLY = IIf(bReturnDirectory, Left$(sFileName, InStrRev(sFileName, "\")), Right$(sFileName, Len(sFileName) - InStrRev(sFileName, "\")))
        '<EhFooter>
        Exit Function

FILE_TITLE_ONLY_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modComDialog.FILE_TITLE_ONLY " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function FILE_EXISTS(sFileName As String) As Boolean
        '<EhHeader>
        On Error GoTo FILE_EXISTS_Err
        '</EhHeader>

100     If sFileName <> "" Then FILE_EXISTS = (Dir(sFileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) <> "")
        '<EhFooter>
        Exit Function

FILE_EXISTS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modComDialog.FILE_EXISTS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function GetFileInName(stitle As String, sFilter As String) As String
        '<EhHeader>
        On Error GoTo GetFileInName_Err
        '</EhHeader>
        Dim Filename As String
100     Filename = FILE_DIALOG(frmCiphers, False, stitle, sFilter)

102     If Filename = "" Then Exit Function
104     If Not FILE_EXISTS(Filename) Then MsgBox Chr$(34) + Filename + Chr$(34) + vbCrLf + "This file does not exist.": Exit Function
106     If FileLen(Filename) = 0 Then MsgBox Chr$(34) + Filename + Chr$(34) + vbCrLf + "File Length is Zero.": Exit Function
108     GetFileInName = Filename
        '<EhFooter>
        Exit Function

GetFileInName_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modComDialog.GetFileInName " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function GetFileOutName(stitle As String, sFilter As String) As String
        '<EhHeader>
        On Error GoTo GetFileOutName_Err
        '</EhHeader>
        Dim Filename As String
100     Filename = FILE_DIALOG(frmCiphers, False, stitle, sFilter)

102     If Filename = "" Then Exit Function
104     GetFileOutName = Filename
        '<EhFooter>
        Exit Function

GetFileOutName_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modComDialog.GetFileOutName " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
